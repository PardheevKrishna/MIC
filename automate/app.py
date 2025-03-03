import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO

from openpyxl import load_workbook  # for preserving all original sheets

# -------------- FIXED EXCEL FILE PATH --------------
FILE_PATH = "input.xlsx"  # <-- Change this to your actual file path

st.set_page_config(page_title="Team Report Dashboard", layout="wide")
st.title("Team Report Dashboard (Fixed Excel File)")

# -------------- GLOBAL ALLOWED VALUES (for categorical columns) --------------
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


def process_excel_file(file_path):
    """
    Reads and processes each employee sheet from 'file_path'.
    Returns two DataFrames:
      - working_hours_details: row-level data with columns including:
          Employee, Status Date (Every Friday), Main project,
          Name of the Project, Start Date, Completion Date (if present),
          Weekly Time Spent(Hrs), Work Hours, PTO Hours, Month (yyyy-mm), WeekFriday (mm-dd-yyyy)
      - violations_df: DataFrame of violations with columns:
          Employee, Violation Type, Violation Details, Location, Violation Date
        where Violation Type is one of:
          "Invalid value", "Working hours less than 40", "Start date change"
    """
    # Read "Home" sheet to get employee names from column F (starting row 3)
    home_df = pd.read_excel(file_path, sheet_name="Home", header=None)
    employee_names = home_df.iloc[2:, 5].dropna().astype(str).tolist()

    xls = pd.ExcelFile(file_path)
    all_sheet_names = xls.sheet_names

    working_hours_details_list = []
    violations_list = []

    # For start date validation: track the first start date per (project, month)
    project_month_info = {}

    # Exceptions for start date checks
    start_date_exceptions = [
        "Internal meetings", "Internal Meetings", "Internal meeting", "internal meeting",
        "External meetings", "External Meeting", "External meeting", "external meetings",
        "Sick leave", "Sick Leave", "Sick day",
        "Annual meeting", "annual meeting", "Traveling", "Develop/Dev training",
        "Internal Taining", "internal taining", "Interview"
    ]

    for emp in employee_names:
        if emp not in all_sheet_names:
            st.warning(f"Sheet for employee '{emp}' not found; skipping.")
            continue

        # Read the employee sheet
        df = pd.read_excel(file_path, sheet_name=emp)
        # Normalize headers: remove newline characters and extra spaces.
        df.columns = [str(c).replace("\n", " ").strip() for c in df.columns]

        required_cols = [
            "Status Date (Every Friday)", "Main project",
            "Name of the Project", "Start Date", "Weekly Time Spent(Hrs)"
        ]
        for rc in required_cols:
            if rc not in df.columns:
                st.error(f"Column '{rc}' not found in sheet '{emp}'.")
                return None, None

        # Convert date columns to datetime
        df["Status Date (Every Friday)"] = pd.to_datetime(df["Status Date (Every Friday)"], errors="coerce")
        df["Start Date"] = pd.to_datetime(df["Start Date"], errors="coerce")
        if "Completion Date" in df.columns:
            df["Completion Date"] = pd.to_datetime(df["Completion Date"], errors="coerce")

        df["Employee"] = emp
        # RowNumber used for updating later
        df["RowNumber"] = df.index + 2  # assume row 1 is header, data starts row 2

        # ---------- 1) ALLOWED VALUES VALIDATION ----------
        for col, allowed_list in allowed_values.items():
            if col not in df.columns:
                continue  # skip if the column doesn't exist
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

        # ---------- 2) START DATE VALIDATION (per project per month) ----------
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
                if key not in project_month_info:
                    project_month_info[key] = {
                        "start_date": start_val,
                        "sheet": emp
                    }
                else:
                    baseline = project_month_info[key]["start_date"]
                    if start_val != baseline:
                        violation_date = row["Status Date (Every Friday)"]
                        baseline_str = baseline.strftime('%m-%d-%Y') if pd.notna(baseline) else 'NaT'
                        current_str = start_val.strftime('%m-%d-%Y') if pd.notna(start_val) else 'NaT'
                        violations_list.append({
                            "Employee": emp,
                            "Violation Type": "Start date change",
                            "Violation Details": (
                                f"Expected {baseline_str}, found {current_str} "
                                f"for '{proj}' in {month_key}"
                            ),
                            "Location": f"Sheet '{emp}', Row {row['RowNumber']}",
                            "Violation Date": violation_date
                        })

        # ---------- 3) WEEKLY HOURS VALIDATION (>= 40) ----------
        df["Weekly Time Spent(Hrs)"] = pd.to_numeric(df["Weekly Time Spent(Hrs)"], errors="coerce").fillna(0)
        # Filter rows where "Status Date (Every Friday)" is a valid Friday
        friday_rows = df[df["Status Date (Every Friday)"].dt.weekday == 4]
        unique_fridays = friday_rows["Status Date (Every Friday)"].dropna().unique()
        for friday in unique_fridays:
            week_start = friday - timedelta(days=4)
            # Check all rows that fall between Monday and Friday of that week
            week_df = df[(df["Status Date (Every Friday)"] >= week_start) &
                         (df["Status Date (Every Friday)"] <= friday)]
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

        # ---------- 4) PREPARE ADDITIONAL COLUMNS ----------
        df["PTO Hours"] = df.apply(
            lambda r: r["Weekly Time Spent(Hrs)"] if "PTO" in str(r["Main project"]) else 0,
            axis=1
        )
        df["Work Hours"] = df.apply(
            lambda r: r["Weekly Time Spent(Hrs)"] if "PTO" not in str(r["Main project"]) else 0,
            axis=1
        )

        # Create "Month" and "WeekFriday" columns
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

                # If any month is selected, break down by week; else aggregate monthly.
                if months_chosen:
                    summary = (
                        filtered_df.groupby(["Employee", "Month", "WeekFriday"], dropna=False)
                        .agg({"Work Hours": "sum", "PTO Hours": "sum"})
                        .reset_index()
                    )
                else:
                    summary = (
                        filtered_df.groupby(["Employee", "Month"], dropna=False)
                        .agg({"Work Hours": "sum", "PTO Hours": "sum"})
                        .reset_index()
                    )

                st.dataframe(summary, use_container_width=True)

                team_buffer = BytesIO()
                with pd.ExcelWriter(team_buffer, engine="openpyxl") as writer:
                    summary.to_excel(writer, sheet_name="Filtered_Team_Monthly", index=False)
                team_buffer.seek(0)
                st.download_button(
                    "Download Filtered Team Monthly",
                    data=team_buffer,
                    file_name="filtered_team_monthly.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
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
                st.download_button(
                    "Download Filtered Working Hours",
                    data=wh_buffer,
                    file_name="filtered_working_hours.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
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

                st.dataframe(filtered_v, use_container_width=True)

                v_buffer = BytesIO()
                with pd.ExcelWriter(v_buffer, engine="openpyxl") as writer:
                    filtered_v.to_excel(writer, sheet_name="Filtered_Violations", index=False)
                v_buffer.seek(0)
                st.download_button(
                    "Download Filtered Violations",
                    data=v_buffer,
                    file_name="filtered_violations.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.info("Select filters and click 'Filter Data' to view results.")

    # -------------- TAB 3: UPDATE DATA (Automatic Mode) --------------
    with tabs[3]:
        st.subheader("Data Update - Automatic Mode")
        st.info(
            "In this mode, for each Main project within a month, the **Start Date** is set to the earliest date (min) "
            "and the **Completion Date** to the latest date (max). For categorical fields, choose whether to update with "
            "the first occurrence or the most frequent value. The updated Excel will contain **all original sheets**, "
            "including 'Home', but each employee sheet will have the updated rows."
        )

        # Filter update tab by Employee and Month
        all_employees = sorted(working_hours_details["Employee"].dropna().unique())
        all_months = sorted(working_hours_details["Month"].dropna().unique())

        with st.form("update_data_automatic_form"):
            col_filter1, col_filter2 = st.columns(2)
            employees_chosen = col_filter1.multiselect("Select Employee(s)", options=all_employees, default=all_employees)
            months_chosen = col_filter2.multiselect("Select Month(s)", options=all_months, default=all_months)

            st.markdown("### Categorical Fields Update Options")
            cat_update_modes = {}
            for col in allowed_values.keys():
                mode_choice = st.radio(
                    f"For column '{col}', choose update mode:",
                    options=["First occurrence", "Most Occurrence"],
                    index=0,
                    key=col
                )
                cat_update_modes[col] = "first" if mode_choice == "First occurrence" else "most"

            update_btn = st.form_submit_button("Update Data Automatically")

        if update_btn:
            # Make a mask for the rows we want to update
            mask = (
                working_hours_details["Employee"].isin(employees_chosen)
                & working_hours_details["Month"].isin(months_chosen)
            )

            updated_data = working_hours_details.copy()
            filtered_df = working_hours_details[mask].copy()

            def update_group(group):
                # Convert to datetime in case something changed
                if "Start Date" in group.columns:
                    group["Start Date"] = pd.to_datetime(group["Start Date"], errors="coerce")
                    group["Start Date"] = group["Start Date"].min()  # earliest

                if "Completion Date" in group.columns:
                    group["Completion Date"] = pd.to_datetime(group["Completion Date"], errors="coerce")
                    group["Completion Date"] = group["Completion Date"].max()  # latest

                # Update each categorical column as per selected mode
                for ccol, mode in cat_update_modes.items():
                    if ccol in group.columns:
                        non_null = group[ccol].dropna()
                        if mode == "first":
                            new_val = non_null.iloc[0] if not non_null.empty else None
                        else:
                            mode_series = non_null.mode()
                            new_val = mode_series.iloc[0] if not mode_series.empty else None
                        group[ccol] = new_val

                return group

            # Group the filtered data by (Employee, Month, Main project) and apply the update
            updated_filtered = filtered_df.groupby(
                ["Employee", "Month", "Main project"], group_keys=False
            ).apply(update_group)

            # Put updated rows back into the full dataset
            updated_data.loc[mask, :] = updated_filtered

            # ==============================
            # WRITE OUT A NEW EXCEL WITH ALL ORIGINAL SHEETS
            # ==============================
            # We'll read the original workbook, and for each sheet:
            #   - If it's an employee sheet, update row-by-row from updated_data
            #   - Otherwise, keep as-is

            # 1) Figure out who the employees are (from "Home" sheet read earlier)
            #    We already have them in 'employee_names' from the process function
            #    But let's ensure we keep that in scope
            home_df = pd.read_excel(FILE_PATH, sheet_name="Home", header=None)
            employee_names = home_df.iloc[2:, 5].dropna().astype(str).tolist()

            # 2) Prepare an ExcelWriter in memory
            output_buffer = BytesIO()
            with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
                # read the original workbook to replicate each sheet
                original_xls = pd.ExcelFile(FILE_PATH)
                for sheet_name in original_xls.sheet_names:
                    df_orig = pd.read_excel(FILE_PATH, sheet_name=sheet_name)

                    if sheet_name in employee_names:
                        # This sheet belongs to an employee, so we do row-by-row updates
                        df_orig["RowNumber"] = df_orig.index + 2  # to match how we assigned in process_excel_file

                        # Get the portion of updated_data for this employee
                        df_emp_updated = updated_data[updated_data["Employee"] == sheet_name].copy()

                        # We'll align rows using "RowNumber"
                        for idx_emp, row_emp in df_emp_updated.iterrows():
                            row_num = row_emp["RowNumber"]
                            mask_row = df_orig["RowNumber"] == row_num

                            # Update only columns that exist in df_orig
                            for col in df_emp_updated.columns:
                                if col in df_orig.columns:
                                    df_orig.loc[mask_row, col] = row_emp[col]

                        # drop RowNumber before saving
                        df_orig.drop(columns=["RowNumber"], inplace=True, errors="ignore")

                        # Write the updated employee sheet
                        df_orig.to_excel(writer, sheet_name=sheet_name, index=False)
                    else:
                        # For non-employee sheets (e.g. "Home"), just write them as is
                        df_orig.to_excel(writer, sheet_name=sheet_name, index=False)

            output_buffer.seek(0)

            st.success("Data updated successfully! All sheets (including 'Home') are preserved in the new Excel.")
            st.info("You can now re-run your validation on this updated Excel if needed.")
            st.download_button(
                "Download Updated Excel (All Sheets)",
                data=output_buffer,
                file_name="updated_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    # -------------- DOWNLOAD FULL REPORT (Original) --------------
    st.markdown("---")
    st.subheader("Download Full Excel Report (Original Data)")
    if not working_hours_details.empty or not violations_df.empty:
        full_buffer = BytesIO()
        with pd.ExcelWriter(full_buffer, engine="openpyxl") as writer:
            working_hours_details.to_excel(writer, sheet_name="Working_Hours_Details", index=False)
            violations_df.to_excel(writer, sheet_name="Violations", index=False)
        full_buffer.seek(0)
        st.download_button(
            "Download Full Excel Report",
            data=full_buffer,
            file_name="full_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )