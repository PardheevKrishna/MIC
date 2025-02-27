import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO

# -------------- FIXED EXCEL FILE PATH --------------
FILE_PATH = "input.xlsx"  # Change to your actual path

st.set_page_config(page_title="Team Report Dashboard", layout="wide")
st.title("Team Report Dashboard (Fixed Excel File)")

# -------------- PROCESS FUNCTION --------------
def process_excel_file(file_path):
    """
    Reads and processes each employee sheet from 'file_path'.
    Returns two main DataFrames:
      - working_hours_details: row-level data for all employees, with columns:
          Employee, RowNumber, Status Date (Every Friday),
          Main project, Name of the Project, Start Date,
          Weekly Time Spent(Hrs), Work Hours, PTO Hours,
          Month (yyyy-mm), WeekFriday (mm-dd-yyyy)
      - violations_df: row of violations with:
          Employee, Violation Type, Violation Details,
          Location, Violation Date
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

    # For start-date checks: track the first start date per (project, month)
    project_month_info = {}

    # Allowed single-token values for the last six columns
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

    # If "Main project" or "Name of the Project" exactly equals one of these, skip start-date check
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
                return None, None

        df["Employee"] = emp
        df["RowNumber"] = df.index + 2  # assume header is row 1

        # Convert Status Date
        df["Status Date (Every Friday)"] = pd.to_datetime(
            df["Status Date (Every Friday)"], format="%m-%d-%Y", errors="coerce"
        )

        # ---------- Allowed Values Check ----------
        for col, allowed_list in allowed_values.items():
            for i, val in df[col].items():
                if pd.isna(val):
                    continue
                tokens = [t.strip() for t in str(val).split(",") if t.strip()]
                if len(tokens) != 1 or tokens[0] not in allowed_list:
                    # For "Invalid value" violation, use the row's date if available
                    violation_date = df.at[i, "Status Date (Every Friday)"]
                    violations_list.append({
                        "Employee": emp,
                        "Violation Type": "Invalid value",
                        "Violation Details": f"Invalid value in '{col}': '{val}'",
                        "Location": f"Sheet '{emp}', Row {df.at[i, 'RowNumber']}",
                        "Violation Date": violation_date  # can be NaT if not available
                    })

        # ---------- Start Date Validation (project+month) ----------
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
                        violation_date = row["Status Date (Every Friday)"]
                        violations_list.append({
                            "Employee": emp,
                            "Violation Type": "Start date change",
                            "Violation Details": (
                                f"Expected {baseline.strftime('%m-%d-%Y') if pd.notna(baseline) else 'NaT'}, "
                                f"found {current_start.strftime('%m-%d-%Y') if pd.notna(current_start) else 'NaT'} "
                                f"for '{proj}' in {month_key}"
                            ),
                            "Location": f"Sheet '{emp}', Row {row['RowNumber']}",
                            "Violation Date": violation_date
                        })

        # ---------- Weekly Hours (≥ 40) ----------
        df["Weekly Time Spent(Hrs)"] = pd.to_numeric(df["Weekly Time Spent(Hrs)"], errors="coerce").fillna(0)
        friday_rows = df[df["Status Date (Every Friday)"].dt.weekday == 4]
        unique_fridays = friday_rows["Status Date (Every Friday)"].dropna().unique()
        for friday in unique_fridays:
            week_start = friday - timedelta(days=4)
            week_df = df[(df["Status Date (Every Friday)"] >= week_start) &
                         (df["Status Date (Every Friday)"] <= friday)]
            total_hrs = week_df["Weekly Time Spent(Hrs)"].sum()
            if total_hrs < 40:
                # For a weekly violation, store the Friday date
                for i2, r2 in week_df.iterrows():
                    pass  # we could list the rownumbers, but we only store one violation
                row_nums = ", ".join(str(x) for x in week_df["RowNumber"].tolist())
                violations_list.append({
                    "Employee": emp,
                    "Violation Type": "Working hours less than 40",
                    "Violation Details": f"Insufficient weekly hours: {total_hrs}",
                    "Location": f"Sheet '{emp}', Rows: {row_nums}",
                    "Violation Date": friday
                })

        # ---------- Prepare row-level PTO/Work, Month, Week ----------
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

        # Store the entire detail
        working_hours_details_list.append(df)

    # Combine all employees
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

    # ========== BUILD TABS FOR DASHBOARD ==========
    tabs = st.tabs(["Team Monthly Summary", "Working Hours Summary", "Violations"])

    # -------------- TAB 1: TEAM MONTHLY SUMMARY --------------
    with tabs[0]:
        st.subheader("Team Monthly Summary")

        if working_hours_details.empty:
            st.info("No data available.")
        else:
            # We'll do an on-the-fly aggregator by month (and employee) AFTER we filter by employee, month, week
            # so we can also incorporate the "Select Week" requirement.
            all_employees = sorted(working_hours_details["Employee"].dropna().unique())
            all_months = sorted(working_hours_details["Month"].dropna().unique())
            all_weeks = sorted(working_hours_details["WeekFriday"].dropna().unique())

            with st.form("monthly_summary_form"):
                c1, c2 = st.columns([0.7, 0.3])
                employees_chosen = c1.multiselect("Select Employee(s)", options=all_employees, default=[])
                select_all_emp = c2.checkbox("Select All Employees")

                c3, c4 = st.columns([0.7, 0.3])
                months_chosen = c3.multiselect("Select Month(s)", options=all_months, default=[])
                select_all_months = c4.checkbox("Select All Months")

                # If at least one month is chosen, only show weeks from those months
                if months_chosen:
                    subset_for_weeks = working_hours_details[working_hours_details["Month"].isin(months_chosen)]
                    possible_weeks = sorted(subset_for_weeks["WeekFriday"].dropna().unique())
                else:
                    possible_weeks = all_weeks

                c5, c6 = st.columns([0.7, 0.3])
                weeks_chosen = c5.multiselect("Select Week(s) (Friday date)", options=possible_weeks, default=[])
                select_all_weeks = c6.checkbox("Select All Weeks")

                filter_btn = st.form_submit_button("Filter Data")

            if filter_btn:
                # Apply the "Select All" logic
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

                # Now do a monthly aggregator
                # Summarize by (Employee, Month) => sum of Work Hours, sum of PTO Hours
                monthly_summary = (
                    filtered_df.groupby(["Employee", "Month"], dropna=False)
                    .agg({"Work Hours": "sum", "PTO Hours": "sum"})
                    .reset_index()
                )

                st.dataframe(monthly_summary, use_container_width=True)

                # Download the aggregator
                team_buffer = BytesIO()
                with pd.ExcelWriter(team_buffer, engine="openpyxl") as writer:
                    monthly_summary.to_excel(writer, sheet_name="Filtered_Team_Monthly", index=False)
                team_buffer.seek(0)
                st.download_button(
                    "Download Filtered Team Monthly",
                    data=team_buffer,
                    file_name="filtered_team_monthly.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.info("Select filters above and click 'Filter Data' to view results.")

    # -------------- TAB 2: WORKING HOURS SUMMARY --------------
    with tabs[1]:
        st.subheader("Working Hours Summary")
        if working_hours_details.empty:
            st.info("No data available.")
        else:
            all_emps_wh = sorted(working_hours_details["Employee"].dropna().unique())
            all_months_wh = sorted(working_hours_details["Month"].dropna().unique())
            all_weeks_wh = sorted(working_hours_details["WeekFriday"].dropna().unique())

            with st.form("working_hours_form"):
                c1, c2 = st.columns([0.7, 0.3])
                emp_chosen_wh = c1.multiselect("Select Employee(s)", options=all_emps_wh, default=[])
                select_all_emp_wh = c2.checkbox("Select All Employees")

                c3, c4 = st.columns([0.7, 0.3])
                months_chosen_wh = c3.multiselect("Select Month(s)", options=all_months_wh, default=[])
                select_all_months_wh = c4.checkbox("Select All Months")

                # If months chosen, only show those weeks
                if months_chosen_wh:
                    subset_for_weeks_wh = working_hours_details[working_hours_details["Month"].isin(months_chosen_wh)]
                    possible_weeks_wh = sorted(subset_for_weeks_wh["WeekFriday"].dropna().unique())
                else:
                    possible_weeks_wh = all_weeks_wh

                c5, c6 = st.columns([0.7, 0.3])
                weeks_chosen_wh = c5.multiselect("Select Week(s) (Friday date)", options=possible_weeks_wh, default=[])
                select_all_weeks_wh = c6.checkbox("Select All Weeks")

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

    # -------------- TAB 3: VIOLATIONS --------------
    with tabs[2]:
        st.subheader("Violations")
        if violations_df.empty:
            st.info("No violations found.")
        else:
            # Distinct employees
            all_emps_v = sorted(violations_df["Employee"].dropna().unique())
            # Distinct violation types
            all_types_v = ["Invalid value", "Working hours less than 40", "Start date change"]
            # (If your data has more, you can expand this list or just read from violations_df.)

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