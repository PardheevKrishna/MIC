import streamlit as st
import pandas as pd
import re
from datetime import datetime, timedelta
from io import BytesIO
from openpyxl import load_workbook

# -------------- FIXED EXCEL FILE PATH --------------
FILE_PATH = "input.xlsx"  # <-- Change this to your actual file path

st.set_page_config(page_title="Team Report Dashboard", layout="wide")
st.title("Team Report Dashboard (Skip Exceptions)")

# -------------- GLOBAL ALLOWED VALUES --------------
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

# -------------- START DATE EXCEPTION KEYWORDS --------------
# Any row containing one of these in either "Main project" or "Name of the Project"
# is fully skipped from start date validation.
exception_keywords = [
    "internal meeting", "internal meetings",
    "external meeting", "external meetings",
    "training", "trainings",
    "sick leave", "sick day",
    "annual meeting", "annual meetings",
    "traveling",
    "develop/dev training",
    "internal taining", "internal training", "internal trainings",
    "interview"
]

def is_exception_row(main_str: str, proj_str: str) -> bool:
    """
    Return True if any exception keyword is found in either main_str or proj_str.
    Both are expected to be lowercased.
    """
    for kw in exception_keywords:
        if kw in main_str or kw in proj_str:
            return True
    return False

def process_excel_file(file_path):
    """
    Reads each employee sheet, skipping start date validation if the row has any exception keyword.
    Returns:
      - working_hours_details: all rows from all employee sheets (with added columns).
      - violations_df: list of all violations (invalid value, working hours < 40, or start date change).
    """
    # Read "Home" sheet to get employee names (column F from row 3 down)
    home_df = pd.read_excel(file_path, sheet_name="Home", header=None)
    employee_names = home_df.iloc[2:, 5].dropna().astype(str).tolist()

    xls = pd.ExcelFile(file_path)
    all_sheet_names = xls.sheet_names

    working_hours_details_list = []
    violations_list = []

    # Dictionary to track the first start date per (normalized main project, normalized project name, month)
    project_month_info = {}

    for emp in employee_names:
        if emp not in all_sheet_names:
            st.warning(f"Sheet for employee '{emp}' not found; skipping.")
            continue

        df = pd.read_excel(file_path, sheet_name=emp)
        df.columns = [str(c).replace("\n", " ").strip() for c in df.columns]

        required_cols = [
            "Status Date (Every Friday)", "Main project",
            "Name of the Project", "Start Date", "Weekly Time Spent(Hrs)"
        ]
        for rc in required_cols:
            if rc not in df.columns:
                st.error(f"Column '{rc}' not found in sheet '{emp}'.")
                return None, None

        # Convert date columns
        df["Status Date (Every Friday)"] = pd.to_datetime(df["Status Date (Every Friday)"], errors="coerce")
        df["Start Date"] = pd.to_datetime(df["Start Date"], errors="coerce")
        if "Completion Date" in df.columns:
            df["Completion Date"] = pd.to_datetime(df["Completion Date"], errors="coerce")

        df["Employee"] = emp
        df["RowNumber"] = df.index + 2

        # --- 1) ALLOWED VALUES VALIDATION ---
        for col, allowed_list in allowed_values.items():
            if col not in df.columns:
                continue
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

        # --- 2) START DATE VALIDATION ---
        for i, row in df.iterrows():
            main_raw = str(row["Main project"]) if pd.notna(row["Main project"]) else ""
            proj_raw = str(row["Name of the Project"]) if pd.notna(row["Name of the Project"]) else ""
            main_norm = main_raw.strip().lower()
            proj_norm = proj_raw.strip().lower()

            start_val = row["Start Date"]
            status_date = row["Status Date (Every Friday)"]
            if pd.isna(status_date):
                continue

            # Skip if this row matches an exception keyword
            if is_exception_row(main_norm, proj_norm):
                # We do not track or violate
                continue

            # Otherwise, proceed
            if proj_norm and pd.notna(start_val):
                month_key = status_date.strftime("%Y-%m")
                key = (main_norm, proj_norm, month_key)
                if key not in project_month_info:
                    project_month_info[key] = start_val
                else:
                    baseline = project_month_info[key]
                    if start_val != baseline:
                        baseline_str = baseline.strftime('%m-%d-%Y') if pd.notna(baseline) else 'NaT'
                        current_str = start_val.strftime('%m-%d-%Y') if pd.notna(start_val) else 'NaT'
                        violations_list.append({
                            "Employee": emp,
                            "Violation Type": "Start date change",
                            "Violation Details": f"Expected {baseline_str}, found {current_str} for '{main_raw}' / '{proj_raw}' in {month_key}",
                            "Location": f"Sheet '{emp}', Row {row['RowNumber']}",
                            "Violation Date": status_date
                        })

        # --- 3) WEEKLY HOURS VALIDATION ---
        df["Weekly Time Spent(Hrs)"] = pd.to_numeric(df["Weekly Time Spent(Hrs)"], errors="coerce").fillna(0)
        friday_rows = df[df["Status Date (Every Friday)"].dt.weekday == 4]
        unique_fridays = friday_rows["Status Date (Every Friday)"].dropna().unique()
        for friday in unique_fridays:
            week_start = friday - timedelta(days=4)
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

        # --- 4) Extra columns ---
        df["PTO Hours"] = df.apply(
            lambda r: r["Weekly Time Spent(Hrs)"] if "PTO" in str(r["Main project"]) else 0,
            axis=1
        )
        df["Work Hours"] = df.apply(
            lambda r: r["Weekly Time Spent(Hrs)"] if "PTO" not in str(r["Main project"]) else 0,
            axis=1
        )
        df["Month"] = df["Status Date (Every Friday)"].dt.to_period("M").astype(str)
        df["WeekFriday"] = df["Status Date (Every Friday)"].dt.strftime("%m-%d-%Y")

        working_hours_details_list.append(df)

    if working_hours_details_list:
        working_hours_details = pd.concat(working_hours_details_list, ignore_index=True)
    else:
        working_hours_details = pd.DataFrame()

    violations_df = pd.DataFrame(violations_list)
    return working_hours_details, violations_df


# -------------- READ & PROCESS --------------
result = process_excel_file(FILE_PATH)
if result is None:
    st.error("Error processing the Excel file.")
else:
    working_hours_details, violations_df = result
    st.success("Reports generated successfully!")

    # -------------- CREATE TABS --------------
    tabs = st.tabs(["Team Monthly Summary", "Working Hours Summary", "Violations", "Update Data"])

    # --- TAB 0: TEAM MONTHLY SUMMARY ---
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

    # --- TAB 1: WORKING HOURS SUMMARY ---
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

    # --- TAB 2: VIOLATIONS ---
    with tabs[2]:
        st.subheader("Violations")
        if violations_df.empty:
            st.info("No violations found.")
        else:
            total_violations = len(violations_df)
            st.markdown(f"**Total Violations: {total_violations}**")

            v_counts = violations_df["Violation Type"].value_counts()
            for vtype in ["Invalid value", "Working hours less than 40", "Start date change"]:
                st.markdown(f"- **{vtype}**: {v_counts.get(vtype, 0)}")

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

    # --- TAB 3: UPDATE DATA (Automatic Mode) ---
    with tabs[3]:
        st.subheader("Data Update - Automatic Mode")
        st.info(
            "Here you can implement your logic to update Start Date, Completion Date, etc. "
            "Rows that match any exception keyword should be fully skipped, just like in the violation check."
        )
        st.write("Implement your update logic here, grouping by the same (normalized) keys. "
                 "After updating, download the new Excel and re-run the code with the new file as input.")