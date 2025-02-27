import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO

# -------------- FIXED EXCEL FILE PATH --------------
FILE_PATH = "input.xlsx"  # <-- Change to your actual path

st.set_page_config(page_title="Team Report Dashboard", layout="wide")
st.title("Team Report Dashboard (Fixed Excel File)")

# -------------- PROCESS FUNCTION --------------
def process_excel_file(file_path):
    """
    Reads and processes each employee sheet from 'file_path'.
    Returns:
      - team_monthly_summary: DataFrame of monthly PTO/Work sums per employee
      - working_hours_details: DataFrame of all rows from all employees
      - violations_df: DataFrame with 'Violation Type' in {Invalid value, Working hours less than 40, Start date change}
    """
    home_df = pd.read_excel(file_path, sheet_name="Home", header=None)
    employee_names = home_df.iloc[2:, 5].dropna().astype(str).tolist()

    xls = pd.ExcelFile(file_path)
    all_sheet_names = xls.sheet_names

    employee_reports_list = []
    working_hours_details_list = []
    violations_list = []

    # For start-date checks: track first start date per (project, month)
    project_month_info = {}

    # Allowed single-token values for the last six columns (exact, case sensitive)
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

    # If "Main project" or "Name of the Project" exactly matches any of these, skip start-date check
    start_date_exceptions = ["Annual Leave"]

    # -------------- PROCESS EACH EMPLOYEE --------------
    for emp in employee_names:
        if emp not in all_sheet_names:
            st.warning(f"Sheet for employee '{emp}' not found; skipping.")
            continue

        df = pd.read_excel(file_path, sheet_name=emp)
        # Normalize headers
        df.columns = [str(c).replace("\n", " ").strip() for c in df.columns]

        required_cols = [
            "Status Date (Every Friday)", "Main project",
            "Name of the Project", "Start Date", "Weekly Time Spent(Hrs)"
        ]
        for rc in required_cols:
            if rc not in df.columns:
                st.error(f"Column '{rc}' not found in sheet '{emp}'.")
                return None

        df["Employee"] = emp
        df["RowNumber"] = df.index + 2  # assume header is row 1

        # Convert Status Date
        df["Status Date (Every Friday)"] = pd.to_datetime(
            df["Status Date (Every Friday)"], format="%m-%d-%Y", errors="coerce"
        )

        # ---------- 1) ALLOWED VALUES ----------
        for col, allowed_list in allowed_values.items():
            for i, val in df[col].items():
                if pd.isna(val):
                    continue
                tokens = [t.strip() for t in str(val).split(",") if t.strip()]
                if len(tokens) != 1 or tokens[0] not in allowed_list:
                    violations_list.append({
                        "Employee": emp,
                        "Violation Type": "Invalid value",
                        "Violation Details": f"Invalid value in '{col}': '{val}'",
                        "Location": f"Sheet '{emp}', Row {df.at[i, 'RowNumber']}"
                    })

        # ---------- 2) START DATE (project+month) ----------
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
                    project_month_info[key] = {
                        "start_date": current_start,
                        "sheet": emp,
                        "row": row["RowNumber"]
                    }
                else:
                    baseline = project_month_info[key]["start_date"]
                    if current_start != baseline:
                        old_str = baseline.strftime("%m-%d-%Y") if pd.notna(baseline) else "NaT"
                        new_str = current_start.strftime("%m-%d-%Y") if pd.notna(current_start) else "NaT"
                        violations_list.append({
                            "Employee": emp,
                            "Violation Type": "Start date change",
                            "Violation Details": f"Expected {old_str}, found {new_str} for '{proj}' in {month_key}",
                            "Location": f"Sheet '{emp}', Row {row['RowNumber']}"
                        })

        # ---------- 3) WEEKLY HOURS ----------
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
                    "Location": f"Sheet '{emp}', Rows: {row_nums}"
                })

        # ---------- 4) MONTHLY SUMMARY (PTO vs Work) ----------
        df["PTO Hours"] = df.apply(
            lambda r: r["Weekly Time Spent(Hrs)"] if "PTO" in str(r["Main project"]) else 0,
            axis=1
        )
        df["Work Hours"] = df.apply(
            lambda r: r["Weekly Time Spent(Hrs)"] if "PTO" not in str(r["Main project"]) else 0,
            axis=1
        )
        df["Month"] = df["Status Date (Every Friday)"].dt.to_period("M").astype(str)

        monthly_summary = df.groupby("Month").agg({
            "Work Hours": "sum",
            "PTO Hours": "sum"
        }).reset_index()
        monthly_summary["Employee"] = emp
        employee_reports_list.append(monthly_summary)

        # Keep entire detail
        working_hours_details_list.append(df)

    # Combine
    if employee_reports_list:
        team_monthly_summary = pd.concat(employee_reports_list, ignore_index=True)
    else:
        team_monthly_summary = pd.DataFrame()

    if working_hours_details_list:
        working_hours_details = pd.concat(working_hours_details_list, ignore_index=True)
    else:
        working_hours_details = pd.DataFrame()

    violations_df = pd.DataFrame(violations_list)
    return team_monthly_summary, working_hours_details, violations_df

# -------------- READ & PROCESS THE EXCEL --------------
result = process_excel_file(FILE_PATH)
if result is None:
    st.error("Error processing the Excel file.")
else:
    team_monthly_summary, working_hours_details, violations_df = result
    st.success("Reports generated successfully!")

    # ========== BUILD TABS FOR DASHBOARD ==========
    tabs = st.tabs(["Team Monthly Summary", "Working Hours Summary", "Violations"])

    # ---------------- TAB 1: TEAM MONTHLY SUMMARY ----------------
    with tabs[0]:
        st.subheader("Team Monthly Summary")
        if team_monthly_summary.empty:
            st.info("No data available.")
        else:
            all_employees = sorted(team_monthly_summary["Employee"].unique())
            all_months = sorted(team_monthly_summary["Month"].unique())

            # Use a form so user can pick multiple options & only see changes after "Filter Data"
            with st.form("monthly_summary_form"):
                # Employees
                c1, c2 = st.columns([0.7, 0.3])
                employees_chosen = c1.multiselect("Select Employee(s)", options=all_employees, default=[])
                select_all_emp = c2.checkbox("Select All Employees")

                # Months
                c3, c4 = st.columns([0.7, 0.3])
                months_chosen = c3.multiselect("Select Month(s)", options=all_months, default=[])
                select_all_months = c4.checkbox("Select All Months")

                filter_btn = st.form_submit_button("Filter Data")

            if filter_btn:
                if select_all_emp:
                    employees_chosen = all_employees
                if select_all_months:
                    months_chosen = all_months

                filtered_team = team_monthly_summary.copy()
                if employees_chosen:
                    filtered_team = filtered_team[filtered_team["Employee"].isin(employees_chosen)]
                if months_chosen:
                    filtered_team = filtered_team[filtered_team["Month"].isin(months_chosen)]

                st.dataframe(filtered_team, use_container_width=True)

                # Download filtered
                team_buffer = BytesIO()
                with pd.ExcelWriter(team_buffer, engine="openpyxl") as writer:
                    filtered_team.to_excel(writer, sheet_name="Filtered_Team_Monthly", index=False)
                team_buffer.seek(0)
                st.download_button(
                    "Download Filtered Team Monthly",
                    data=team_buffer,
                    file_name="filtered_team_monthly.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.info("Select filters above and click 'Filter Data' to view results.")

    # ---------------- TAB 2: WORKING HOURS SUMMARY ----------------
    with tabs[1]:
        st.subheader("Working Hours Summary")
        if working_hours_details.empty:
            st.info("No data available.")
        else:
            # Precompute Month & Week columns
            working_hours_details["Month"] = working_hours_details["Status Date (Every Friday)"].dt.to_period("M").astype(str)
            working_hours_details["WeekFriday"] = working_hours_details["Status Date (Every Friday)"].dt.strftime("%m-%d-%Y")

            all_emps_wh = sorted(working_hours_details["Employee"].unique())
            all_months_wh = sorted(working_hours_details["Month"].unique())
            all_weeks = sorted(working_hours_details["WeekFriday"].dropna().unique())

            with st.form("working_hours_form"):
                # Employees
                c1, c2 = st.columns([0.7, 0.3])
                emp_chosen_wh = c1.multiselect("Select Employee(s)", options=all_emps_wh, default=[])
                select_all_emp_wh = c2.checkbox("Select All Employees")

                # Months
                c3, c4 = st.columns([0.7, 0.3])
                months_chosen_wh = c3.multiselect("Select Month(s)", options=all_months_wh, default=[])
                select_all_months_wh = c4.checkbox("Select All Months")

                # Weeks
                c5, c6 = st.columns([0.7, 0.3])
                weeks_chosen_wh = c5.multiselect("Select Week(s) (Friday date)", options=all_weeks, default=[])
                select_all_weeks_wh = c6.checkbox("Select All Weeks")

                filter_btn_wh = st.form_submit_button("Filter Data")

            if filter_btn_wh:
                if select_all_emp_wh:
                    emp_chosen_wh = all_emps_wh
                if select_all_months_wh:
                    months_chosen_wh = all_months_wh
                if select_all_weeks_wh:
                    weeks_chosen_wh = all_weeks

                filtered_wh = working_hours_details.copy()
                if emp_chosen_wh:
                    filtered_wh = filtered_wh[filtered_wh["Employee"].isin(emp_chosen_wh)]
                if months_chosen_wh:
                    filtered_wh = filtered_wh[filtered_wh["Month"].isin(months_chosen_wh)]
                if weeks_chosen_wh:
                    filtered_wh = filtered_wh[filtered_wh["WeekFriday"].isin(weeks_chosen_wh)]

                st.dataframe(filtered_wh, use_container_width=True)

                # Download filtered
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
                st.info("Select filters above and click 'Filter Data' to view results.")

    # ---------------- TAB 3: VIOLATIONS ----------------
    with tabs[2]:
        st.subheader("Violations")
        if violations_df.empty:
            st.info("No violations found.")
        else:
            # The user wants exactly 3 possible types:
            #   "Invalid value", "Working hours less than 40", "Start date change"
            # We'll display only those. If your data has others, they'd appear anyway, but let's assume it's just these 3.
            all_emps_v = sorted(violations_df["Employee"].dropna().unique())
            all_types_v = sorted(violations_df["Violation Type"].dropna().unique())

            with st.form("violations_form"):
                c1, c2 = st.columns([0.7, 0.3])
                emp_chosen_v = c1.multiselect("Select Employee(s)", options=all_emps_v, default=[])
                select_all_emp_v = c2.checkbox("Select All Employees")

                c3, c4 = st.columns([0.7, 0.3])
                types_chosen_v = c3.multiselect("Select Violation Type(s)", options=all_types_v, default=[])
                select_all_types_v = c4.checkbox("Select All Types")

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

                # Download filtered
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
                st.info("Select filters above and click 'Filter Data' to view results.")

    # -------------- DOWNLOAD FULL REPORT --------------
    st.markdown("---")
    st.subheader("Download Full Excel Report")
    if not team_monthly_summary.empty or not working_hours_details.empty or not violations_df.empty:
        full_buffer = BytesIO()
        with pd.ExcelWriter(full_buffer, engine="openpyxl") as writer:
            team_monthly_summary.to_excel(writer, sheet_name="Team_Monthly_Summary", index=False)
            working_hours_details.to_excel(writer, sheet_name="Working_Hours_Details", index=False)
            violations_df.to_excel(writer, sheet_name="Violations", index=False)
        full_buffer.seek(0)

        st.download_button(
            "Download Full Excel Report",
            data=full_buffer,
            file_name="full_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )