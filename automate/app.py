import streamlit as st
import pandas as pd
from datetime import timedelta
from io import BytesIO
from openpyxl import load_workbook

# ---------- CONFIGURATION ----------
FILE_PATH = "input.xlsx"  # CHANGE this to your actual Excel file path

# -------------- CUSTOM CSS & TITLE --------------
st.markdown("""
<style>
.mySkipButton {
    visibility: hidden;
}
</style>
""", unsafe_allow_html=True)

st.title("Team Report Dashboard")

# ===================== PROCESS FUNCTION =====================
def process_excel_file(file_path):
    """
    Reads each employee sheet (employee names from "Home") and returns two DataFrames:
      - working_details: row-level data for all employees
      - violations_df: flagged violations (Invalid value, Start date change, Weekly < 40, etc.)
    """
    # Allowed values for validation
    allowed_values = {
        "Functional Area (CRIT, CRIT - Data Management, CRIT - Data Governance, CRIT - Regulatory Reporting, CRIT - Portfolio Reporting, CRIT - Transformation)": [
            "CRIT", "CRIT - Data Management", "CRIT - Data Governance", "CRIT - Regulatory Reporting", "CRIT - Portfolio Reporting", "CRIT - Transformation"
        ],
        "Project Category (Data Infrastructure, Monitoring & Insights, Analytics / Strategy Development, GDA Related, Trainings and Team Meeting)": [
            "Data Infrastructure", "Monitoring & Insights", "Analytics / Strategy Development", "GDA Related", "Trainings and Team Meeting"
        ],
        "Complexity (H,M,L)": ["H", "M", "L"],
        "Novelity (BAU repetitive, One time repetitive, New one time)": [
            "BAU repetitive", "One time repetitive", "New one time"
        ],
        "Output Type (Core production work, Ad-hoc long-term projects, Ad-hoc short-term projects, Business Management, Administration, Trainings/L&D activities, Others) :": [
            "Core production work", "Ad-hoc long-term projects", "Ad-hoc short-term projects", "Business Management", "Administration", "Trainings/L&D activities", "Others"
        ],
        "Impact type (Customer Experience, Financial impact, Insights, Risk reduction, Others)": [
            "Customer Experience", "Financial impact", "Insights", "Risk reduction", "Others"
        ]
    }

    # Exceptions for start date check
    start_date_exceptions = [
        "Internal meetings", "Internal Meetings", "Internal meeting", "internal meeting",
        "External meetings", "External Meeting", "External meeting", "external meetings",
        "Sick leave", "Sick Leave", "Sick day",
        "Annual meeting", "annual meeting", "Traveling", "Develop/Dev training",
        "Internal Taining", "internal taining", "Interview"
    ]

    try:
        home_df = pd.read_excel(file_path, sheet_name="Home", header=None)
    except Exception as e:
        st.error(f"Error reading Home sheet: {e}")
        return None, None

    employee_names = home_df.iloc[2:, 5].dropna().astype(str).tolist()
    xls = pd.ExcelFile(file_path)
    all_sheet_names = xls.sheet_names

    working_list = []
    viol_list = []  # list to collect violations
    project_month_info = {}

    for emp in employee_names:
        if emp not in all_sheet_names:
            continue
        try:
            df = pd.read_excel(file_path, sheet_name=emp)
        except Exception as e:
            st.warning(f"Could not read sheet for {emp}: {e}")
            continue

        df.columns = [str(c).replace("\n", " ").strip() for c in df.columns]
        required_cols = ["Status Date (Every Friday)", "Main project", "Name of the Project", "Start Date", "Weekly Time Spent(Hrs)"]
        if not all(c in df.columns for c in required_cols):
            continue

        df["Employee"] = emp
        df["RowNumber"] = df.index + 2  # Excel row number (header row = 1)
        df["Status Date (Every Friday)"] = pd.to_datetime(df["Status Date (Every Friday)"], format="%m-%d-%Y", errors="coerce")

        # (1) Validate allowed values
        for col, a_list in allowed_values.items():
            if col not in df.columns:
                continue
            for i, val in df[col].items():
                if pd.isna(val):
                    continue
                tokens = [t.strip() for t in str(val).split(",") if t.strip()]
                if len(tokens) != 1 or tokens[0] not in a_list:
                    viol_list.append({
                        "Employee": emp,
                        "Violation Type": "Invalid value",
                        "Violation Details": f"{col} = {val}",
                        "Location": f"Sheet {emp}, Row {df.at[i, 'RowNumber']}",
                        "Violation Date": df.at[i, "Status Date (Every Friday)"]
                    })

        # (2) Check start date consistency (unless project is in exceptions)
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
                    project_month_info[key] = current_start
                else:
                    baseline = project_month_info[key]
                    if pd.notna(current_start) and pd.notna(baseline) and current_start != baseline:
                        old_str = baseline.strftime("%m-%d-%Y")
                        new_str = current_start.strftime("%m-%d-%Y")
                        viol_list.append({
                            "Employee": emp,
                            "Violation Type": "Start date change",
                            "Violation Details": f"{proj}: expected {old_str}, got {new_str}",
                            "Location": f"Sheet {emp}, Row {row['RowNumber']}",
                            "Violation Date": row["Status Date (Every Friday)"]
                        })

        # (3) Weekly hours check
        df["Weekly Time Spent(Hrs)"] = pd.to_numeric(df["Weekly Time Spent(Hrs)"], errors="coerce").fillna(0)
        friday_dates = df[(df["Status Date (Every Friday)"].dt.weekday == 4) & (df["Status Date (Every Friday)"].notna())]["Status Date (Every Friday)"].unique()
        for friday in friday_dates:
            if pd.isna(friday):
                continue
            friday_str = friday.strftime("%m-%d-%Y")
            week_start = friday - timedelta(days=4)
            week_df = df[(df["Status Date (Every Friday)"] >= week_start) & (df["Status Date (Every Friday)"] <= friday)]
            total_hrs = week_df["Weekly Time Spent(Hrs)"].sum()
            if total_hrs < 40:
                row_nums_str = ", ".join(str(x) for x in week_df["RowNumber"].tolist())
                viol_list.append({
                    "Employee": emp,
                    "Violation Type": "Working hours less than 40",
                    "Violation Details": f"Week ending {friday_str} insufficient hours",
                    "Location": f"Sheet {emp}, Rows: {row_nums_str}",
                    "Violation Date": friday
                })

        # (4) Additional columns for further processing
        df["PTO Hours"] = df.apply(lambda r: r["Weekly Time Spent(Hrs)"] if "PTO" in str(r["Main project"]) else 0, axis=1)
        df["Work Hours"] = df.apply(lambda r: r["Weekly Time Spent(Hrs)"] if "PTO" not in str(r["Main project"]) else 0, axis=1)
        df["Month"] = df["Status Date (Every Friday)"].dt.to_period("M").astype(str)
        df["WeekFriday"] = df["Status Date (Every Friday)"].dt.strftime("%m-%d-%Y").fillna("N/A")
        df["UniqueID"] = df["Employee"] + "_" + df["RowNumber"].astype(str)

        working_list.append(df)

    working_details = pd.concat(working_list, ignore_index=True) if working_list else pd.DataFrame()
    violations_df = pd.DataFrame(viol_list)
    return working_details, violations_df

# ======= SESSION STATE DATA LOADING / INITIALIZATION =======
if "working_details" not in st.session_state or "violations_df" not in st.session_state:
    wd, vd = process_excel_file(FILE_PATH)
    if wd is None or vd is None:
        st.error("Error processing the Excel file.")
        st.stop()
    st.session_state["working_details"] = wd
    st.session_state["violations_df"] = vd
    st.success("Reports generated successfully!")
else:
    st.success("Using cached data from session state.")

working_details = st.session_state["working_details"]
violations_df = st.session_state["violations_df"]

# ========== TABS ==========
tab1, tab2, tab3 = st.tabs(["Team Monthly Summary", "Working Hours Summary", "Violations and Update"])

# ========== TAB 1: TEAM MONTHLY SUMMARY ==========
with tab1:
    st.subheader("Team Monthly Summary")
    if working_details.empty:
        st.info("No data available.")
    else:
        all_emps = sorted(working_details["Employee"].dropna().unique())
        all_months = sorted(working_details["Month"].dropna().unique())
        all_weeks = sorted(working_details["WeekFriday"].dropna().unique())
        with st.form("tm_form"):
            c1, c2 = st.columns([0.7, 0.3])
            emp_sel = c1.multiselect("Select Employee(s)", options=all_emps)
            all_emp_check = c2.checkbox("Select All Employees", key="tm_all_emp")
            c3, c4 = st.columns([0.7, 0.3])
            month_sel = c3.multiselect("Select Month(s)", options=all_months)
            all_month_check = c4.checkbox("Select All Months", key="tm_all_month")
            possible_weeks = sorted(working_details[working_details["Month"].isin(month_sel)]["WeekFriday"].dropna().unique()) if month_sel else all_weeks
            c5, c6 = st.columns([0.7, 0.3])
            week_sel = c5.multiselect("Select Week(s)", options=possible_weeks)
            all_week_check = c6.checkbox("Select All Weeks", key="tm_all_week")
            submit_tm = st.form_submit_button("Filter Data")
        if submit_tm:
            if all_emp_check:
                emp_sel = all_emps
            if all_month_check:
                month_sel = all_months
            if all_week_check:
                week_sel = possible_weeks
            df_tm = working_details.copy()
            if emp_sel:
                df_tm = df_tm[df_tm["Employee"].isin(emp_sel)]
            if month_sel:
                df_tm = df_tm[df_tm["Month"].isin(month_sel)]
            if week_sel:
                df_tm = df_tm[df_tm["WeekFriday"].isin(week_sel)]
            if month_sel:
                summary = df_tm.groupby(["Employee", "Month", "WeekFriday"]).agg({"Work Hours": "sum", "PTO Hours": "sum"}).reset_index()
            else:
                summary = df_tm.groupby(["Employee", "Month"]).agg({"Work Hours": "sum", "PTO Hours": "sum"}).reset_index()
            st.dataframe(summary, use_container_width=True)
            buf = BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                summary.to_excel(writer, sheet_name="Team_Monthly", index=False)
            buf.seek(0)
            st.download_button("Download Team Monthly Summary", data=buf,
                               file_name="Team_Monthly_Summary.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ========== TAB 2: WORKING HOURS SUMMARY ==========
with tab2:
    st.subheader("Working Hours Summary")
    if working_details.empty:
        st.info("No data available.")
    else:
        all_emps_wh = sorted(working_details["Employee"].dropna().unique())
        all_months_wh = sorted(working_details["Month"].dropna().unique())
        all_weeks_wh = sorted(working_details["WeekFriday"].dropna().unique())
        with st.form("wh_form"):
            col1, col2 = st.columns([0.7, 0.3])
            emp_sel_wh = col1.multiselect("Select Employee(s)", options=all_emps_wh)
            all_emp_wh = col2.checkbox("Select All Employees", key="wh_all_emp")
            col3, col4 = st.columns([0.7, 0.3])
            month_sel_wh = col3.multiselect("Select Month(s)", options=all_months_wh)
            all_month_wh = col4.checkbox("Select All Months", key="wh_all_month")
            poss_weeks_wh = sorted(working_details[working_details["Month"].isin(month_sel_wh)]["WeekFriday"].dropna().unique()) if month_sel_wh else all_weeks_wh
            col5, col6 = st.columns([0.7, 0.3])
            week_sel_wh = col5.multiselect("Select Week(s)", options=poss_weeks_wh)
            all_week_wh = col6.checkbox("Select All Weeks", key="wh_all_week")
            submit_wh = st.form_submit_button("Filter Data")
        if submit_wh:
            if all_emp_wh:
                emp_sel_wh = all_emps_wh
            if all_month_wh:
                month_sel_wh = all_months_wh
            if all_week_wh:
                week_sel_wh = poss_weeks_wh
            df_wh = working_details.copy()
            if emp_sel_wh:
                df_wh = df_wh[df_wh["Employee"].isin(emp_sel_wh)]
            if month_sel_wh:
                df_wh = df_wh[df_wh["Month"].isin(month_sel_wh)]
            if week_sel_wh:
                df_wh = df_wh[df_wh["WeekFriday"].isin(week_sel_wh)]
            st.dataframe(df_wh, use_container_width=True)
            buf_wh = BytesIO()
            with pd.ExcelWriter(buf_wh, engine="openpyxl") as writer:
                df_wh.to_excel(writer, sheet_name="Working_Hours", index=False)
            buf_wh.seek(0)
            st.download_button("Download Working Hours", data=buf_wh,
                               file_name="Working_Hours_Summary.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ========== TAB 3: VIOLATIONS & UPDATE ==========
with tab3:
    st.subheader("Violations and Update")
    if violations_df.empty:
        st.info("No violations found.")
    else:
        # Step 1: Filter Violations
        all_emps_v = sorted(violations_df["Employee"].dropna().unique())
        all_types_v = ["Invalid value", "Working hours less than 40", "Start date change"]

        with st.form("violations_filter_form"):
            col1_v, col2_v = st.columns([0.7, 0.3])
            emp_sel_v = col1_v.multiselect("Select Employee(s)", options=all_emps_v)
            select_all_emp_v = col2_v.checkbox("Select All Employees")
            col3_v, col4_v = st.columns([0.7, 0.3])
            type_sel_v = col3_v.multiselect("Select Violation Type(s)", options=all_types_v)
            select_all_type_v = col4_v.checkbox("Select All Violation Types")
            filter_btn_v = st.form_submit_button("Filter Violations")
        if filter_btn_v:
            if select_all_emp_v:
                emp_sel_v = all_emps_v
            if select_all_type_v:
                type_sel_v = all_types_v
            df_v = violations_df.copy()
            if emp_sel_v:
                df_v = df_v[df_v["Employee"].isin(emp_sel_v)]
            if type_sel_v:
                df_v = df_v[df_v["Violation Type"].isin(type_sel_v)]
            st.dataframe(df_v, use_container_width=True)

            # Step 2: Select rows for update (by UniqueID)
            all_ids = sorted(df_v.apply(lambda r: f"{r['Employee']}_{r['Location'].split('Row ')[-1]}", axis=1).unique())
            st.markdown("#### Select Rows to Update (by UniqueID)")
            select_all_rows = st.checkbox("Select All Rows")
            selected_ids = all_ids if select_all_rows else st.multiselect("Select UniqueIDs", options=all_ids)

            # Step 3: Choose update mode
            update_mode = st.radio("Select Update Mode", options=["Automatic", "Manual"], index=0)
            proceed_btn = st.button("Load Selected Rows for Editing")
            if proceed_btn:
                if not selected_ids:
                    st.error("No rows selected for update.")
                else:
                    st.session_state["selected_rows"] = selected_ids
                    st.markdown(f"**Rows selected for update:** {selected_ids}")

                    st.markdown("### Edit Each Selected Row Below")
                    updated_data = {}  # dictionary to collect updated row info

                    # Function to compute auto suggestions for automatic mode
                    def compute_auto_suggestions(row, df):
                        proj_group = df[df["Main project"] == row["Main project"]]
                        auto_start = pd.to_datetime(proj_group["Start Date"], errors="coerce").min()
                        auto_comp = None
                        if "Completion Date" in proj_group.columns:
                            auto_comp = pd.to_datetime(proj_group["Completion Date"], errors="coerce").max()
                        cat_fields = [
                            "Functional Area (CRIT, CRIT - Data Management, CRIT - Data Governance, CRIT - Regulatory Reporting, CRIT - Portfolio Reporting, CRIT - Transformation)",
                            "Project Category (Data Infrastructure, Monitoring & Insights, Analytics / Strategy Development, GDA Related, Trainings and Team Meeting)",
                            "Complexity (H,M,L)",
                            "Novelity (BAU repetitive, One time repetitive, New one time)",
                            "Output Type (Core production work, Ad-hoc long-term projects, Ad-hoc short-term projects, Business Management, Administration, Trainings/L&D activities, Others) :",
                            "Impact type (Customer Experience, Financial impact, Insights, Risk reduction, Others)"
                        ]
                        cat_suggestions = {}
                        for cf in cat_fields:
                            if cf in proj_group.columns and not proj_group[cf].dropna().empty:
                                cat_suggestions[cf] = proj_group[cf].mode().iloc[0]
                            else:
                                cat_suggestions[cf] = ""
                        auto_start_str = auto_start.strftime("%m-%d-%Y") if pd.notna(auto_start) else ""
                        auto_comp_str = auto_comp.strftime("%m-%d-%Y") if auto_comp is not None and pd.notna(auto_comp) else ""
                        return auto_start_str, auto_comp_str, cat_suggestions

                    cat_fields = [
                        "Functional Area (CRIT, CRIT - Data Management, CRIT - Data Governance, CRIT - Regulatory Reporting, CRIT - Portfolio Reporting, CRIT - Transformation)",
                        "Project Category (Data Infrastructure, Monitoring & Insights, Analytics / Strategy Development, GDA Related, Trainings and Team Meeting)",
                        "Complexity (H,M,L)",
                        "Novelity (BAU repetitive, One time repetitive, New one time)",
                        "Output Type (Core production work, Ad-hoc long-term projects, Ad-hoc short-term projects, Business Management, Administration, Trainings/L&D activities, Others) :",
                        "Impact type (Customer Experience, Financial impact, Insights, Risk reduction, Others)"
                    ]

                    # Build a lookup from UniqueID to the corresponding row in working_details
                    working_details_dict = {}
                    for idx, r in working_details.iterrows():
                        uid = r["UniqueID"]
                        working_details_dict[uid] = r

                    for uid in selected_ids:
                        if uid not in working_details_dict:
                            st.warning(f"No data found for {uid}")
                            continue
                        row = working_details_dict[uid]
                        with st.expander(f"Edit row {uid}", expanded=True):
                            if update_mode == "Automatic":
                                auto_start_str, auto_comp_str, cat_sugg = compute_auto_suggestions(row, working_details)
                                st.write(f"Main project: {row['Main project']}")
                                new_start = st.text_input("Start Date", value=auto_start_str, key=f"{uid}_start")
                                new_comp = ""
                                if "Completion Date" in row and "Completion Date" in working_details.columns:
                                    new_comp = st.text_input("Completion Date", value=auto_comp_str, key=f"{uid}_comp")
                                cat_vals = {}
                                for cf in cat_fields:
                                    cat_vals[cf] = st.text_input(cf, value=cat_sugg[cf], key=f"{uid}_{cf}")
                                updated_data[uid] = {
                                    "Employee": row["Employee"],
                                    "RowNumber": row["RowNumber"],
                                    "Start Date": new_start,
                                    "Completion Date": new_comp,
                                    **cat_vals
                                }
                            else:
                                current_start = str(row.get("Start Date", ""))
                                new_start = st.text_input("Start Date", value=current_start, key=f"{uid}_start")
                                current_comp = ""
                                if "Completion Date" in row and "Completion Date" in working_details.columns:
                                    current_comp = str(row.get("Completion Date", ""))
                                    new_comp = st.text_input("Completion Date", value=current_comp, key=f"{uid}_comp")
                                else:
                                    new_comp = ""
                                cat_vals = {}
                                for cf in cat_fields:
                                    cat_current = str(row.get(cf, ""))
                                    cat_vals[cf] = st.text_input(cf, value=cat_current, key=f"{uid}_{cf}")
                                updated_data[uid] = {
                                    "Employee": row["Employee"],
                                    "RowNumber": row["RowNumber"],
                                    "Start Date": new_start,
                                    "Completion Date": new_comp,
                                    **cat_vals
                                }
                    # Once editing is complete, update when button is clicked
                    if st.button("Update Excel"):
                        try:
                            wb = load_workbook(FILE_PATH)
                        except Exception as e:
                            st.error(f"Error opening workbook: {e}")
                            st.stop()

                        for uid, row_vals in updated_data.items():
                            sheet_name = row_vals["Employee"]
                            if sheet_name not in wb.sheetnames:
                                continue
                            ws = wb[sheet_name]
                            headers = {cell.value: cell.column for cell in ws[1]}
                            r_num = row_vals["RowNumber"]
                            if "Start Date" in headers and row_vals["Start Date"]:
                                ws.cell(row=r_num, column=headers["Start Date"], value=row_vals["Start Date"])
                            if "Completion Date" in headers and row_vals["Completion Date"]:
                                ws.cell(row=r_num, column=headers["Completion Date"], value=row_vals["Completion Date"])
                            for cf in cat_fields:
                                if cf in headers and row_vals.get(cf, ""):
                                    ws.cell(row=r_num, column=headers[cf], value=row_vals[cf])
                        try:
                            wb.save(FILE_PATH)
                            st.success("Excel file updated successfully.")
                        except Exception as e:
                            st.error(f"Error saving workbook: {e}")

                        # --- Workaround: Update the session state working_details in memory ---
                        for uid, row_vals in updated_data.items():
                            mask = st.session_state["working_details"]["UniqueID"] == uid
                            if mask.any():
                                idx = st.session_state["working_details"].index[mask][0]
                                for col, new_val in row_vals.items():
                                    st.session_state["working_details"].at[idx, col] = new_val
                        st.success("Session data updated. The changes will now appear in your dashboard.")
        else:
            st.info("Please use the form above to filter violations.")