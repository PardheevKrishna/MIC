import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO
import os

# ------------- FIXED EXCEL FILE PATH -------------
FILE_PATH = "input.xlsx"  # Change this to your actual path

# ------------- PAGE CONFIG & TITLE -------------
st.set_page_config(page_title="Team Report Dashboard", layout="wide")
st.title("Team Report Dashboard (Fixed Excel File)")

# ------------- HELPER FUNCTION -------------
def process_excel_file(file_path):
    # Read the "Home" sheet to get employee names from column F (starting row 3).
    home_df = pd.read_excel(file_path, sheet_name="Home", header=None)
    employee_names = home_df.iloc[2:, 5].dropna().astype(str).tolist()
    
    xls = pd.ExcelFile(file_path)
    all_sheet_names = xls.sheet_names

    # Prepare containers for final DataFrames
    employee_reports_list = []
    working_hours_details_list = []
    violations_list = []

    # Track the first start date encountered per (project, month)
    project_month_info = {}

    # Allowed values for the last six columns
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

    # If either "Main project" or "Name of the Project" exactly equals one of these, skip start-date check
    start_date_exceptions = ["Annual Leave"]

    # ---------- Process Each Employee Sheet ----------
    for emp in employee_names:
        if emp not in all_sheet_names:
            # If the sheet doesn't exist, just warn and skip
            st.warning(f"No sheet found for employee '{emp}'. Skipping.")
            continue
        
        df = pd.read_excel(file_path, sheet_name=emp)
        # Normalize column headers
        df.columns = [str(c).replace("\n", " ").strip() for c in df.columns]
        
        # Required columns
        req_cols = [
            "Status Date (Every Friday)", "Main project",
            "Name of the Project", "Start Date", "Weekly Time Spent(Hrs)"
        ]
        for rc in req_cols:
            if rc not in df.columns:
                st.error(f"Column '{rc}' not found in sheet '{emp}'.")
                return None
        
        # Add columns for reference
        df["Employee"] = emp
        df["RowNumber"] = df.index + 2  # Assuming header is row 1

        # Convert Status Date to datetime
        df["Status Date (Every Friday)"] = pd.to_datetime(
            df["Status Date (Every Friday)"], format='%m-%d-%Y', errors='coerce'
        )

        # ---------- Allowed Values Validation ----------
        for col, allowed in allowed_values.items():
            for i, val in df[col].items():
                if pd.isna(val):
                    continue
                # Split on commas, strip whitespace
                tokens = [t.strip() for t in str(val).split(",") if t.strip()]
                if len(tokens) != 1 or tokens[0] not in allowed:
                    violations_list.append({
                        "Employee": emp,
                        "Violation Type": f"Invalid value in '{col}': '{val}'",
                        "Location": f"Sheet '{emp}', Row {df.at[i, 'RowNumber']}"
                    })

        # ---------- Start Date Validation (per project per month) ----------
        for i, row in df.iterrows():
            proj = row["Name of the Project"]
            start_val = row["Start Date"]
            mp_val = str(row["Main project"]).strip() if pd.notna(row["Main project"]) else ""
            proj_val = str(proj).strip() if pd.notna(proj) else ""
            
            # Skip if in exceptions
            if mp_val in start_date_exceptions or proj_val in start_date_exceptions:
                continue
            
            if pd.notna(proj) and pd.notna(start_val) and pd.notna(row["Status Date (Every Friday)"]):
                month_key = row["Status Date (Every Friday)"].strftime("%Y-%m")
                key = (proj, month_key)
                current_start = pd.to_datetime(start_val, format='%m-%d-%Y', errors='coerce')
                
                if key not in project_month_info:
                    project_month_info[key] = {
                        "start_date": current_start,
                        "sheet": emp,
                        "row": row["RowNumber"]
                    }
                else:
                    baseline = project_month_info[key]["start_date"]
                    if current_start != baseline:
                        violations_list.append({
                            "Employee": emp,
                            "Violation Type": (
                                f"Start date changed for project '{proj}' in {month_key}: "
                                f"expected {baseline.strftime('%m-%d-%Y') if pd.notna(baseline) else 'NaT'}, "
                                f"found {current_start.strftime('%m-%d-%Y') if pd.notna(current_start) else 'NaT'}"
                            ),
                            "Location": f"Sheet '{emp}', Row {row['RowNumber']}"
                        })
        
        # ---------- Weekly Hours Validation ----------
        df["Weekly Time Spent(Hrs)"] = pd.to_numeric(df["Weekly Time Spent(Hrs)"], errors='coerce').fillna(0)
        fridays = df[df["Status Date (Every Friday)"].dt.weekday == 4]
        unique_fridays = fridays["Status Date (Every Friday)"].dropna().unique()
        
        for friday in unique_fridays:
            week_start = friday - timedelta(days=4)
            week_df = df[(df["Status Date (Every Friday)"] >= week_start) &
                         (df["Status Date (Every Friday)"] <= friday)]
            total_hrs = week_df["Weekly Time Spent(Hrs)"].sum()
            if total_hrs < 40:
                row_nums = ", ".join(str(x) for x in week_df["RowNumber"].tolist())
                violations_list.append({
                    "Employee": emp,
                    "Violation Type": (
                        f"Insufficient weekly work hours: {total_hrs} (<40) for week ending {friday.strftime('%m-%d-%Y')}"
                    ),
                    "Location": f"Sheet '{emp}', Rows: {row_nums}"
                })
        
        # ---------- Monthly Summary (PTO vs. Work) ----------
        df["PTO Hours"] = df.apply(
            lambda r: r["Weekly Time Spent(Hrs)"] if "PTO" in str(r["Main project"]) else 0, axis=1
        )
        df["Work Hours"] = df.apply(
            lambda r: r["Weekly Time Spent(Hrs)"] if "PTO" not in str(r["Main project"]) else 0, axis=1
        )
        df["Month"] = df["Status Date (Every Friday)"].dt.to_period("M").astype(str)
        monthly = df.groupby("Month").agg({"Work Hours": "sum", "PTO Hours": "sum"}).reset_index()
        monthly["Employee"] = emp
        employee_reports_list.append(monthly)

        # Keep track of full working hours detail
        working_hours_details_list.append(df)
    
    # Combine data
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

# ------------- PROCESS THE FIXED EXCEL -------------
result = process_excel_file(FILE_PATH)
if result is None:
    st.error("Error processing the Excel file.")
else:
    team_monthly_summary, working_hours_details, violations_df = result
    st.success("Reports generated successfully!")

    # ========== BUILD DASHBOARD WITH TABS ==========
    tabs = st.tabs(["Team Monthly Summary", "Working Hours Summary", "Violations"])

    # 1) TEAM MONTHLY SUMMARY TAB
    with tabs[0]:
        st.subheader("Team Monthly Summary")

        # -- MULTISELECT: EMPLOYEES (start empty)
        all_employees = sorted(team_monthly_summary["Employee"].unique()) if not team_monthly_summary.empty else []
        if "employee_filter_monthly" not in st.session_state:
            st.session_state["employee_filter_monthly"] = []
        emp_filter = st.multiselect(
            "Select Employee(s)",
            options=all_employees,
            default=st.session_state["employee_filter_monthly"],
            key="employee_filter_monthly"
        )
        # 'Select All' button
        if st.button("Select All Employees", key="select_all_emp_monthly"):
            st.session_state["employee_filter_monthly"] = all_employees
            st.experimental_rerun()

        # -- MULTISELECT: MONTHS (start empty)
        all_months = sorted(team_monthly_summary["Month"].unique()) if not team_monthly_summary.empty else []
        if "month_filter_monthly" not in st.session_state:
            st.session_state["month_filter_monthly"] = []
        month_filter = st.multiselect(
            "Select Month(s)",
            options=all_months,
            default=st.session_state["month_filter_monthly"],
            key="month_filter_monthly"
        )
        if st.button("Select All Months", key="select_all_months_monthly"):
            st.session_state["month_filter_monthly"] = all_months
            st.experimental_rerun()

        filtered_team = team_monthly_summary.copy()
        if emp_filter:
            filtered_team = filtered_team[filtered_team["Employee"].isin(emp_filter)]
        if month_filter:
            filtered_team = filtered_team[filtered_team["Month"].isin(month_filter)]
        
        st.dataframe(filtered_team, use_container_width=True)

        # Download filtered data
        team_buffer = BytesIO()
        with pd.ExcelWriter(team_buffer, engine="openpyxl") as writer:
            filtered_team.to_excel(writer, sheet_name="Filtered_Team_Monthly", index=False)
        team_buffer.seek(0)
        st.download_button(
            "Download Filtered Team Monthly",
            data=team_buffer,
            file_name="filtered_team_monthly.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_filtered_team_monthly"
        )

    # 2) WORKING HOURS SUMMARY TAB
    with tabs[1]:
        st.subheader("Working Hours Summary")

        if not working_hours_details.empty:
            working_hours_details["Month"] = working_hours_details["Status Date (Every Friday)"].dt.to_period("M").astype(str)
            # Distinct weeks = distinct Fridays (non-null)
            working_hours_details["WeekFriday"] = working_hours_details["Status Date (Every Friday)"].dt.strftime("%m-%d-%Y")
        else:
            working_hours_details["Month"] = []
            working_hours_details["WeekFriday"] = []

        # -- EMPLOYEES
        all_emps_wh = sorted(working_hours_details["Employee"].unique()) if not working_hours_details.empty else []
        if "employee_filter_wh" not in st.session_state:
            st.session_state["employee_filter_wh"] = []
        emp_filter_wh = st.multiselect(
            "Select Employee(s)",
            options=all_emps_wh,
            default=st.session_state["employee_filter_wh"],
            key="employee_filter_wh"
        )
        if st.button("Select All Employees", key="select_all_emp_wh"):
            st.session_state["employee_filter_wh"] = all_emps_wh
            st.experimental_rerun()

        # -- MONTHS
        all_months_wh = sorted(working_hours_details["Month"].unique()) if not working_hours_details.empty else []
        if "month_filter_wh" not in st.session_state:
            st.session_state["month_filter_wh"] = []
        month_filter_wh = st.multiselect(
            "Select Month(s)",
            options=all_months_wh,
            default=st.session_state["month_filter_wh"],
            key="month_filter_wh"
        )
        if st.button("Select All Months", key="select_all_months_wh"):
            st.session_state["month_filter_wh"] = all_months_wh
            st.experimental_rerun()

        # -- WEEKS (distinct Friday dates)
        all_weeks = sorted(working_hours_details["WeekFriday"].dropna().unique()) if not working_hours_details.empty else []
        if "week_filter_wh" not in st.session_state:
            st.session_state["week_filter_wh"] = []
        week_filter_wh = st.multiselect(
            "Select Week(s) (Friday date)",
            options=all_weeks,
            default=st.session_state["week_filter_wh"],
            key="week_filter_wh"
        )
        if st.button("Select All Weeks", key="select_all_weeks_wh"):
            st.session_state["week_filter_wh"] = all_weeks
            st.experimental_rerun()

        filtered_wh = working_hours_details.copy()
        if emp_filter_wh:
            filtered_wh = filtered_wh[filtered_wh["Employee"].isin(emp_filter_wh)]
        if month_filter_wh:
            filtered_wh = filtered_wh[filtered_wh["Month"].isin(month_filter_wh)]
        if week_filter_wh:
            filtered_wh = filtered_wh[filtered_wh["WeekFriday"].isin(week_filter_wh)]
        
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
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_filtered_working_hours"
        )

    # 3) VIOLATIONS TAB
    with tabs[2]:
        st.subheader("Violations")
        if violations_df.empty:
            st.info("No violations found.")
        else:
            # Distinct employees in violations
            all_emps_v = sorted(violations_df["Employee"].dropna().unique())
            if "employee_filter_v" not in st.session_state:
                st.session_state["employee_filter_v"] = []
            emp_filter_v = st.multiselect(
                "Select Employee(s)",
                options=all_emps_v,
                default=st.session_state["employee_filter_v"],
                key="employee_filter_v"
            )
            if st.button("Select All Employees", key="select_all_emp_v"):
                st.session_state["employee_filter_v"] = all_emps_v
                st.experimental_rerun()

            # Distinct violation types
            all_types_v = sorted(violations_df["Violation Type"].dropna().unique())
            if "type_filter_v" not in st.session_state:
                st.session_state["type_filter_v"] = []
            type_filter_v = st.multiselect(
                "Select Violation Type(s)",
                options=all_types_v,
                default=st.session_state["type_filter_v"],
                key="type_filter_v"
            )
            if st.button("Select All Types", key="select_all_types_v"):
                st.session_state["type_filter_v"] = all_types_v
                st.experimental_rerun()

            filtered_v = violations_df.copy()
            if emp_filter_v:
                filtered_v = filtered_v[filtered_v["Employee"].isin(emp_filter_v)]
            if type_filter_v:
                filtered_v = filtered_v[filtered_v["Violation Type"].isin(type_filter_v)]
            
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
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_filtered_violations"
            )

    # ------------- Download Full Report -------------
    st.markdown("---")
    st.subheader("Download Full Report")
    # Rebuild the full Excel in memory
    full_buffer = BytesIO()
    with pd.ExcelWriter(full_buffer, engine="openpyxl") as writer:
        team_monthly_summary.to_excel(writer, sheet_name="Team_Monthly_Summary", index=False)
        if not working_hours_details.empty:
            working_hours_details.to_excel(writer, sheet_name="Working_Hours_Details", index=False)
        if not violations_df.empty:
            violations_df.to_excel(writer, sheet_name="Violations", index=False)
    full_buffer.seek(0)

    st.download_button(
        "Download Full Excel Report",
        data=full_buffer,
        file_name="full_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_full_report"
    )