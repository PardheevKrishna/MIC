import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO

# -------------- FIXED EXCEL FILE PATH --------------
FILE_PATH = "input.xlsx"  # CHANGE this to your actual file path

st.set_page_config(page_title="Team Report Dashboard", layout="wide")
st.title("Team Report Dashboard (Fixed Excel File)")

# ===================== PROCESS FUNCTION (for Report Tabs) =====================
def process_excel_file(file_path):
    """
    Reads and processes each employee sheet (listed in the "Home" sheet) from the Excel file.
    Returns two DataFrames:
      - working_hours_details: row-level data from all employee sheets with additional columns:
          Employee, RowNumber, Month (yyyy-mm), WeekFriday (mm-dd-yyyy)
      - violations_df: DataFrame of violations with columns:
          Employee, Violation Type, Violation Details, Location, Violation Date
        where Violation Type is one of:
          "Invalid value", "Working hours less than 40", "Start date change"
    """
    home_df = pd.read_excel(file_path, sheet_name="Home", header=None)
    employee_names = home_df.iloc[2:, 5].dropna().astype(str).tolist()
    xls = pd.ExcelFile(file_path)
    all_sheet_names = xls.sheet_names

    working_hours_details_list = []
    violations_list = []
    project_month_info = {}  # for start date validation

    # Allowed values for six categorical columns (must be a single token exactly matching)
    allowed_values = {
        "Functional Area (CRIT, CRIT - Data Management, CRIT - Data Governance, CRIT - Regulatory Reporting, CRIT - Portfolio Reporting, CRIT - Transformation)":
            ["CRIT", "CRIT - Data Management", "CRIT - Data Governance", "CRIT - Regulatory Reporting", "CRIT - Portfolio Reporting", "CRIT - Transformation"],
        "Project Category (Data Infrastructure, Monitoring & Insights, Analytics / Strategy Development, GDA Related, Trainings and Team Meeting)":
            ["Data Infrastructure", "Monitoring & Insights", "Analytics / Strategy Development", "GDA Related", "Trainings and Team Meeting"],
        "Complexity (H,M,L)":
            ["H", "M", "L"],
        "Novelity (BAU repetitive, One time repetitive, New one time)":
            ["BAU repetitive", "One time repetitive", "New one time"],
        "Output Type (Core production work, Ad-hoc long-term projects, Ad-hoc short-term projects, Business Management, Administration, Trainings/L&D activities, Others) :":
            ["Core production work", "Ad-hoc long-term projects", "Ad-hoc short-term projects", "Business Management", "Administration", "Trainings/L&D activities", "Others"],
        "Impact type (Customer Experience, Financial impact, Insights, Risk reduction, Others)":
            ["Customer Experience", "Financial impact", "Insights", "Risk reduction", "Others"]
    }
    # Exception: if "Main project" or "Name of the Project" equals "Annual Leave", skip start date check.
    start_date_exceptions = ["Annual Leave"]

    for emp in employee_names:
        if emp not in all_sheet_names:
            st.warning(f"Sheet for employee '{emp}' not found; skipping.")
            continue
        df = pd.read_excel(file_path, sheet_name=emp)
        df.columns = [str(c).replace("\n", " ").strip() for c in df.columns]
        required_cols = ["Status Date (Every Friday)", "Main project", "Name of the Project", "Start Date", "Weekly Time Spent(Hrs)"]
        for rc in required_cols:
            if rc not in df.columns:
                st.error(f"Column '{rc}' not found in sheet '{emp}'.")
                return None, None

        df["Employee"] = emp
        df["RowNumber"] = df.index + 2  # assume header is row 1
        df["Status Date (Every Friday)"] = pd.to_datetime(df["Status Date (Every Friday)"], format="%m-%d-%Y", errors="coerce")

        # (1) Allowed values validation
        for col, allowed_list in allowed_values.items():
            for i, val in df[col].items():
                if pd.isna(val):
                    continue
                tokens = [t.strip() for t in str(val).split(",") if t.strip()]
                if len(tokens) != 1 or tokens[0] not in allowed_list:
                    violation_date = df.at[i, "Status Date (Every Friday)"]
                    violations_list.append({
                        "Employee": emp,
                        "Violation Type": "Invalid value",
                        "Violation Details": f"Invalid value in '{col}': '{val}'",
                        "Location": f"Sheet '{emp}', Row {df.at[i, 'RowNumber']}",
                        "Violation Date": violation_date
                    })

        # (2) Start Date consistency per (project, month)
        for i, row in df.iterrows():
            proj = row["Name of the Project"]
            start_val = row["Start Date"]
            mp_val = str(row["Main project"]).strip() if pd.notna(row["Main project"]) else ""
            proj_val = str(proj).strip() if pd.notna(proj) else ""
            if mp_val in start_date_exceptions or proj_val in start_date_exceptions:
                continue
            if pd.notna(proj) and pd.notna(start_val) and pd.notna(row["Status Date (Every Friday)"]):
                month_key = row["Status Date (Every Friday)"].strftime("%Y-%m")
                key = (proj, month_key)
                current_start = pd.to_datetime(start_val, format="%m-%d-%Y", errors="coerce")
                if key not in project_month_info:
                    project_month_info[key] = {"start_date": current_start, "sheet": emp, "row": row["RowNumber"]}
                else:
                    baseline = project_month_info[key]["start_date"]
                    if current_start != baseline:
                        violation_date = row["Status Date (Every Friday)"]
                        violations_list.append({
                            "Employee": emp,
                            "Violation Type": "Start date change",
                            "Violation Details": f"Expected {baseline.strftime('%m-%d-%Y') if pd.notna(baseline) else 'NaT'}, found {current_start.strftime('%m-%d-%Y') if pd.notna(current_start) else 'NaT'} for '{proj}' in {month_key}",
                            "Location": f"Sheet '{emp}', Row {row['RowNumber']}",
                            "Violation Date": violation_date
                        })

        # (3) Weekly hours validation (≥ 40)
        df["Weekly Time Spent(Hrs)"] = pd.to_numeric(df["Weekly Time Spent(Hrs)"], errors="coerce").fillna(0)
        friday_rows = df[df["Status Date (Every Friday)"].dt.weekday == 4]
        unique_fridays = friday_rows["Status Date (Every Friday)"].dropna().unique()
        for friday in unique_fridays:
            week_start = friday - timedelta(days=4)
            week_df = df[(df["Status Date (Every Friday)"] >= week_start) & (df["Status Date (Every Friday)"] <= friday)]
            total_hrs = week_df["Weekly Time Spent(Hrs)"].sum()
            if total_hrs < 40:
                row_nums = ", ".join(str(x) for x in week_df["RowNumber"].tolist())
                violations_list.append({
                    "Employee": emp,
                    "Violation Type": "Working hours less than 40",
                    "Violation Details": f"Insufficient weekly hours: {total_hrs}",
                    "Location": f"Sheet '{emp}', Rows: {row_nums}",
                    "Violation Date": friday
                })

        # (4) Additional computed columns
        df["PTO Hours"] = df.apply(lambda r: r["Weekly Time Spent(Hrs)"] if "PTO" in str(r["Main project"]) else 0, axis=1)
        df["Work Hours"] = df.apply(lambda r: r["Weekly Time Spent(Hrs)"] if "PTO" not in str(r["Main project"]) else 0, axis=1)
        df["Month"] = df["Status Date (Every Friday)"].dt.to_period("M").astype(str)
        df["WeekFriday"] = df["Status Date (Every Friday)"].dt.strftime("%m-%d-%Y")

        working_hours_details_list.append(df)

    if working_hours_details_list:
        working_hours_details = pd.concat(working_hours_details_list, ignore_index=True)
    else:
        working_hours_details = pd.DataFrame()
    violations_df = pd.DataFrame(violations_list)
    return working_hours_details, violations_df

# -------------- READ & PROCESS THE EXCEL --------------
result = process_excel_file(FILE_PATH)
if result is None:
    st.error("Error processing the Excel file.")
else:
    working_hours_details, violations_df = result
    st.success("Reports generated successfully!")

    # ========== DASHBOARD TABS ==========
    tabs = st.tabs(["Team Monthly Summary", "Working Hours Summary", "Violations", "Update Data"])

    # -------------- TAB 0: TEAM MONTHLY SUMMARY --------------
    with tabs[0]:
        st.subheader("Team Monthly Summary")
        if working_hours_details.empty:
            st.info("No data available.")
        else:
            all_employees = sorted(working_hours_details["Employee"].dropna().unique())
            all_months = sorted(working_hours_details["Month"].dropna().unique())
            all_weeks = sorted(working_hours_details["WeekFriday"].dropna().unique())
            with st.form("monthly_summary_form"):
                col_emp, col_emp_select = st.columns([0.7, 0.3])
                employees_chosen = col_emp.multiselect("Select Employee(s)", options=all_employees, default=[])
                select_all_emp = col_emp_select.checkbox("Select All Employees", key="team_emp_all")
                col_month, col_month_select = st.columns([0.7, 0.3])
                months_chosen = col_month.multiselect("Select Month(s)", options=all_months, default=[])
                select_all_months = col_month_select.checkbox("Select All Months", key="team_month_all")
                if months_chosen:
                    subset_weeks = working_hours_details[working_hours_details["Month"].isin(months_chosen)]
                    possible_weeks = sorted(subset_weeks["WeekFriday"].dropna().unique())
                else:
                    possible_weeks = all_weeks
                col_week, col_week_select = st.columns([0.7, 0.3])
                weeks_chosen = col_week.multiselect("Select Week(s) (Friday date)", options=possible_weeks, default=[])
                select_all_weeks = col_week_select.checkbox("Select All Weeks", key="team_week_all")
                filter_btn = st.form_submit_button("Filter Data")
            if filter_btn:
                if select_all_emp:
                    employees_chosen = all_employees
                if select_all_months:
                    months_chosen = all_months
                if select_all_weeks:
                    weeks_chosen = possible_weeks
                filtered_df = working_hours_details.copy()
                if employees_chosen:
                    filtered_df = filtered_df[filtered_df["Employee"].isin(employees_chosen)]
                if months_chosen:
                    filtered_df = filtered_df[filtered_df["Month"].isin(months_chosen)]
                if weeks_chosen:
                    filtered_df = filtered_df[filtered_df["WeekFriday"].isin(weeks_chosen)]
                if months_chosen:
                    summary = filtered_df.groupby(["Employee", "Month", "WeekFriday"], dropna=False).agg({"Work Hours": "sum", "PTO Hours": "sum"}).reset_index()
                else:
                    summary = filtered_df.groupby(["Employee", "Month"], dropna=False).agg({"Work Hours": "sum", "PTO Hours": "sum"}).reset_index()
                st.dataframe(summary, use_container_width=True)
                team_buffer = BytesIO()
                with pd.ExcelWriter(team_buffer, engine="openpyxl") as writer:
                    summary.to_excel(writer, sheet_name="Filtered_Team_Monthly", index=False)
                team_buffer.seek(0)
                st.download_button("Download Filtered Team Monthly", data=team_buffer, file_name="filtered_team_monthly.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.info("Select filters and click 'Filter Data' to view results.")

    # -------------- TAB 1: WORKING HOURS SUMMARY --------------
    with tabs[1]:
        st.subheader("Working Hours Summary")
        if working_hours_details.empty:
            st.info("No data available.")
        else:
            all_emps_wh = sorted(working_hours_details["Employee"].dropna().unique())
            all_months_wh = sorted(working_hours_details["Month"].dropna().unique())
            all_weeks_wh = sorted(working_hours_details["WeekFriday"].dropna().unique())
            with st.form("working_hours_form"):
                col1, col2 = st.columns([0.7, 0.3])
                emp_chosen_wh = col1.multiselect("Select Employee(s)", options=all_emps_wh, default=[])
                select_all_emp_wh = col2.checkbox("Select All Employees", key="wh_emp_all")
                col3, col4 = st.columns([0.7, 0.3])
                months_chosen_wh = col3.multiselect("Select Month(s)", options=all_months_wh, default=[])
                select_all_months_wh = col4.checkbox("Select All Months", key="wh_month_all")
                if months_chosen_wh:
                    subset_weeks_wh = working_hours_details[working_hours_details["Month"].isin(months_chosen_wh)]
                    possible_weeks_wh = sorted(subset_weeks_wh["WeekFriday"].dropna().unique())
                else:
                    possible_weeks_wh = all_weeks_wh
                col5, col6 = st.columns([0.7, 0.3])
                weeks_chosen_wh = col5.multiselect("Select Week(s) (Friday date)", options=possible_weeks_wh, default=[])
                select_all_weeks_wh = col6.checkbox("Select All Weeks", key="wh_week_all")
                filter_btn_wh = st.form_submit_button("Filter Data")
            if filter_btn_wh:
                if select_all_emp_wh:
                    emp_chosen_wh = all_emps_wh
                if select_all_months_wh:
                    months_chosen_wh = all_months_wh
                if select_all_weeks_wh:
                    weeks_chosen_wh = possible_weeks_wh
                filtered_wh = working_hours_details.copy()
                if emp_chosen_wh:
                    filtered_wh = filtered_wh[filtered_wh["Employee"].isin(emp_chosen_wh)]
                if months_chosen_wh:
                    filtered_wh = filtered_wh[filtered_wh["Month"].isin(months_chosen_wh)]
                if weeks_chosen_wh:
                    filtered_wh = filtered_wh[filtered_wh["WeekFriday"].isin(weeks_chosen_wh)]
                st.dataframe(filtered_wh, use_container_width=True)
                wh_buffer = BytesIO()
                with pd.ExcelWriter(wh_buffer, engine="openpyxl") as writer:
                    filtered_wh.to_excel(writer, sheet_name="Filtered_Working_Hours", index=False)
                wh_buffer.seek(0)
                st.download_button("Download Filtered Working Hours", data=wh_buffer, file_name="filtered_working_hours.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.info("Select filters and click 'Filter Data' to view results.")

    # -------------- TAB 2: VIOLATIONS --------------
    with tabs[2]:
        st.subheader("Violations")
        if violations_df.empty:
            st.info("No violations found.")
        else:
            all_emps_v = sorted(violations_df["Employee"].dropna().unique())
            all_types_v = ["Invalid value", "Working hours less than 40", "Start date change"]
            with st.form("violations_form"):
                col1_v, col2_v = st.columns([0.7, 0.3])
                emp_chosen_v = col1_v.multiselect("Select Employee(s)", options=all_emps_v, default=[])
                select_all_emp_v = col2_v.checkbox("Select All Employees", key="v_emp_all")
                col3_v, col4_v = st.columns([0.7, 0.3])
                types_chosen_v = col3_v.multiselect("Select Violation Type(s)", options=all_types_v, default=[])
                select_all_types_v = col4_v.checkbox("Select All Types", key="v_type_all")
                filter_btn_v = st.form_submit_button("Filter Data")
            if filter_btn_v:
                if select_all_emp_v:
                    emp_chosen_v = all_emps_v
                if select_all_types_v:
                    types_chosen_v = all_types_v
                filtered_v = violations_df.copy()
                if emp_chosen_v:
                    filtered_v = filtered_v[filtered_v["Employee"].isin(emp_chosen_v)]
                if types_chosen_v:
                    filtered_v = filtered_v[filtered_v["Violation Type"].isin(types_chosen_v)]
                # Create UniqueID from Employee and RowNumber extracted from Location.
                def extract_row(loc_str):
                    try:
                        return loc_str.split("Row ")[-1]
                    except:
                        return ""
                filtered_v["UniqueID"] = filtered_v.apply(lambda r: f"{r['Employee']}_{extract_row(r['Location'])}", axis=1)
                st.dataframe(filtered_v, use_container_width=True)
                v_buffer = BytesIO()
                with pd.ExcelWriter(v_buffer, engine="openpyxl") as writer:
                    filtered_v.to_excel(writer, sheet_name="Filtered_Violations", index=False)
                v_buffer.seek(0)
                st.download_button("Download Filtered Violations", data=v_buffer, file_name="filtered_violations.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.info("Select filters and click 'Filter Data' to view results.")

    # -------------- TAB 3: UPDATE DATA --------------
    with tabs[3]:
        st.subheader("Update Data for Violations")
        st.markdown("Use the filters below (same as in the Violations tab) to select the rows you wish to update. Then, update the inline checkboxes to select specific rows, choose update options, and click the update button.")
        if violations_df.empty:
            st.info("No violations available for update.")
        else:
            # Ensure UniqueID exists
            def extract_row(loc_str):
                try:
                    return loc_str.split("Row ")[-1]
                except:
                    return ""
            violations_df["UniqueID"] = violations_df.apply(lambda r: f"{r['Employee']}_{extract_row(r['Location'])}", axis=1)
            all_emps_update = sorted(violations_df["Employee"].dropna().unique())
            all_types_update = ["Invalid value", "Working hours less than 40", "Start date change"]
            with st.form("update_filter_form"):
                col1_u, col2_u = st.columns([0.7, 0.3])
                emp_chosen_update = col1_u.multiselect("Select Employee(s)", options=all_emps_update, default=[])
                select_all_emp_update = col2_u.checkbox("Select All Employees", key="update_emp_all")
                col3_u, col4_u = st.columns([0.7, 0.3])
                type_chosen_update = col3_u.multiselect("Select Violation Type(s)", options=all_types_update, default=[])
                select_all_type_update = col4_u.checkbox("Select All Types", key="update_type_all")
                filter_update_btn = st.form_submit_button("Apply Filters")
            if filter_update_btn:
                if select_all_emp_update:
                    emp_chosen_update = all_emps_update
                if select_all_type_update:
                    type_chosen_update = all_types_update
                filtered_update_df = violations_df.copy()
                if emp_chosen_update:
                    filtered_update_df = filtered_update_df[filtered_update_df["Employee"].isin(emp_chosen_update)]
                if type_chosen_update:
                    filtered_update_df = filtered_update_df[filtered_update_df["Violation Type"].isin(type_chosen_update)]
                st.markdown("#### Filtered Violations for Update")
                st.dataframe(filtered_update_df, use_container_width=True)

                # Add an inline "Select" column to allow row-by-row selection.
                if "selected_rows" not in st.session_state:
                    st.session_state["selected_rows"] = []
                # Prepopulate "Select" based on previously selected IDs.
                filtered_update_df["Select"] = filtered_update_df["UniqueID"].apply(lambda uid: uid in st.session_state["selected_rows"])
                updated_table = st.data_editor(filtered_update_df, key="update_data_editor", num_rows="dynamic", use_container_width=True)
                # Save updated selection in session state.
                selected_ids = updated_table[updated_table["Select"] == True]["UniqueID"].tolist()
                st.session_state["selected_rows"] = selected_ids
                st.markdown(f"**Rows selected for update:** {st.session_state['selected_rows']}")

                st.markdown("### Update Options")
                upd_mode = st.radio("Select Update Mode", options=["Automatic", "Manual"], index=0, key="upd_mode")
                categorical_fields = [
                    "Functional Area (CRIT, CRIT - Data Management, CRIT - Data Governance, CRIT - Regulatory Reporting, CRIT - Portfolio Reporting, CRIT - Transformation)",
                    "Project Category (Data Infrastructure, Monitoring & Insights, Analytics / Strategy Development, GDA Related, Trainings and Team Meeting)",
                    "Complexity (H,M,L)",
                    "Novelity (BAU repetitive, One time repetitive, New one time)",
                    "Output Type (Core production work, Ad-hoc long-term projects, Ad-hoc short-term projects, Business Management, Administration, Trainings/L&D activities, Others) :",
                    "Impact type (Customer Experience, Financial impact, Insights, Risk reduction, Others)"
                ]
                # Get suggestions from working_hours_details for the selected rows.
                working_hours_details["UniqueID"] = working_hours_details.apply(lambda r: f"{r['Employee']}_{r['RowNumber']}", axis=1)
                selected_rows_df = working_hours_details[working_hours_details["UniqueID"].isin(st.session_state["selected_rows"])]
                st.markdown("#### Selected Rows Preview")
                st.dataframe(selected_rows_df, use_container_width=True)
                
                if upd_mode == "Automatic":
                    st.markdown("**Automatic Mode**")
                    auto_start_date = selected_rows_df["Start Date"].apply(lambda x: pd.to_datetime(x, format="%m-%d-%Y", errors="coerce")).min()
                    if "Completion Date" in selected_rows_df.columns:
                        auto_completion_date = selected_rows_df["Completion Date"].apply(lambda x: pd.to_datetime(x, format="%m-%d-%Y", errors="coerce")).max()
                    else:
                        auto_completion_date = None
                    st.write("Start Date will be updated to (first occurrence):", auto_start_date.strftime("%m-%d-%Y") if pd.notna(auto_start_date) else "N/A")
                    st.write("Completion Date will be updated to (last occurrence):", auto_completion_date.strftime("%m-%d-%Y") if auto_completion_date is not None and pd.notna(auto_completion_date) else "N/A")
                    auto_choices = {}
                    auto_suggestions = {}
                    for field in categorical_fields:
                        if field in selected_rows_df.columns and not selected_rows_df[field].dropna().empty:
                            first_occurrence = selected_rows_df[field].iloc[0]
                            most_freq = selected_rows_df[field].mode().iloc[0]
                        else:
                            first_occurrence = None
                            most_freq = None
                        auto_suggestions[field] = {"First Occurrence": first_occurrence, "Most Frequent": most_freq}
                        auto_choices[field] = st.radio(f"Select update method for **{field}**", options=["First Occurrence", "Most Frequent"], index=0, key=field+"_upd_auto")
                    upd_btn = st.button("Update Data (Automatic)")
                    if upd_btn:
                        sheets_dict = pd.read_excel(FILE_PATH, sheet_name=None)
                        for sheet_name, df in sheets_dict.items():
                            if sheet_name == "Home":
                                continue
                            df.columns = [str(c).replace("\n", " ").strip() for c in df.columns]
                            if "Status Date (Every Friday)" not in df.columns or "Main project" not in df.columns:
                                continue
                            df["Status Date (Every Friday)"] = pd.to_datetime(df["Status Date (Every Friday)"], format="%m-%d-%Y", errors="coerce")
                            df["RowNumber"] = df.index + 2
                            df["UniqueID"] = df.apply(lambda r: f"{r['Employee']}_{r['RowNumber']}", axis=1)
                            condition = df["UniqueID"].isin(st.session_state["selected_rows"])
                            if condition.any():
                                new_start = auto_start_date.strftime("%m-%d-%Y") if pd.notna(auto_start_date) else None
                                new_completion = auto_completion_date.strftime("%m-%d-%Y") if auto_completion_date is not None and pd.notna(auto_completion_date) else None
                                df.loc[condition, "Start Date"] = new_start if new_start is not None else df.loc[condition, "Start Date"]
                                if "Completion Date" in df.columns and new_completion is not None:
                                    df.loc[condition, "Completion Date"] = new_completion
                                for field in categorical_fields:
                                    if field in df.columns:
                                        chosen_val = auto_suggestions[field][auto_choices[field]]
                                        df.loc[condition, field] = chosen_val
                                sheets_dict[sheet_name] = df
                        with pd.ExcelWriter(FILE_PATH, engine="openpyxl", mode="w") as writer:
                            for sheet_name, df in sheets_dict.items():
                                df.to_excel(writer, sheet_name=sheet_name, index=False)
                        st.success("Selected data updated successfully (Automatic Mode).")
                else:
                    st.markdown("**Manual Mode**")
                    manual_start_date = st.date_input("Select Start Date", value=auto_start_date.date() if pd.notna(auto_start_date) else None)
                    if "Completion Date" in selected_rows_df.columns and auto_completion_date is not None and pd.notna(auto_completion_date):
                        manual_completion_date = st.date_input("Select Completion Date", value=auto_completion_date.date())
                    else:
                        manual_completion_date = None
                    manual_choices = {}
                    allowed_values_manual = {
                        "Functional Area (CRIT, CRIT - Data Management, CRIT - Data Governance, CRIT - Regulatory Reporting, CRIT - Portfolio Reporting, CRIT - Transformation)":
                            ["CRIT", "CRIT - Data Management", "CRIT - Data Governance", "CRIT - Regulatory Reporting", "CRIT - Portfolio Reporting", "CRIT - Transformation"],
                        "Project Category (Data Infrastructure, Monitoring & Insights, Analytics / Strategy Development, GDA Related, Trainings and Team Meeting)":
                            ["Data Infrastructure", "Monitoring & Insights", "Analytics / Strategy Development", "GDA Related", "Trainings and Team Meeting"],
                        "Complexity (H,M,L)":
                            ["H", "M", "L"],
                        "Novelity (BAU repetitive, One time repetitive, New one time)":
                            ["BAU repetitive", "One time repetitive", "New one time"],
                        "Output Type (Core production work, Ad-hoc long-term projects, Ad-hoc short-term projects, Business Management, Administration, Trainings/L&D activities, Others) :":
                            ["Core production work", "Ad-hoc long-term projects", "Ad-hoc short-term projects", "Business Management", "Administration", "Trainings/L&D activities", "Others"],
                        "Impact type (Customer Experience, Financial impact, Insights, Risk reduction, Others)":
                            ["Customer Experience", "Financial impact", "Insights", "Risk reduction", "Others"]
                    }
                    for field in categorical_fields:
                        if field in selected_rows_df.columns:
                            manual_choices[field] = st.selectbox(f"Select value for **{field}**", options=allowed_values_manual[field], key=field+"_upd_manual")
                        else:
                            manual_choices[field] = None
                    upd_btn_manual = st.button("Update Data (Manual)")
                    if upd_btn_manual:
                        sheets_dict = pd.read_excel(FILE_PATH, sheet_name=None)
                        for sheet_name, df in sheets_dict.items():
                            if sheet_name == "Home":
                                continue
                            df.columns = [str(c).replace("\n", " ").strip() for c in df.columns]
                            if "Status Date (Every Friday)" not in df.columns or "Main project" not in df.columns:
                                continue
                            df["Status Date (Every Friday)"] = pd.to_datetime(df["Status Date (Every Friday)"], format="%m-%d-%Y", errors="coerce")
                            df["RowNumber"] = df.index + 2
                            df["UniqueID"] = df.apply(lambda r: f"{r['Employee']}_{r['RowNumber']}", axis=1)
                            condition = df["UniqueID"].isin(st.session_state["selected_rows"])
                            if condition.any():
                                new_start = manual_start_date.strftime("%m-%d-%Y")
                                new_completion = manual_completion_date.strftime("%m-%d-%Y") if manual_completion_date is not None else None
                                df.loc[condition, "Start Date"] = new_start
                                if "Completion Date" in df.columns and new_completion is not None:
                                    df.loc[condition, "Completion Date"] = new_completion
                                for field in categorical_fields:
                                    if field in df.columns:
                                        df.loc[condition, field] = manual_choices[field]
                                sheets_dict[sheet_name] = df
                        with pd.ExcelWriter(FILE_PATH, engine="openpyxl", mode="w") as writer:
                            for sheet_name, df in sheets_dict.items():
                                df.to_excel(writer, sheet_name=sheet_name, index=False)
                        st.success("Selected data updated successfully (Manual Mode).")