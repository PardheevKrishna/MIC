import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment

# ----------------- CONFIGURATION -----------------
FILE_PATH = "input.xlsx"  # Path to the Excel file on disk

st.set_page_config(page_title="Team Report Dashboard", layout="wide")
st.title("Team Report Dashboard")
st.markdown("### Generating Reports from Fixed Excel File")

# ----------------- HELPER FUNCTION: PROCESS EXCEL FILE -----------------
def process_excel_file(file_path):
    # Read the Home sheet (assumes employee names are in column F starting at row 3)
    home_df = pd.read_excel(file_path, sheet_name="Home", header=None)
    employee_names = home_df.iloc[2:, 5].dropna().astype(str).tolist()
    
    xls = pd.ExcelFile(file_path)
    all_sheet_names = xls.sheet_names

    employee_reports_list = []      # Monthly summary per employee (with an "Employee" column)
    working_hours_details_list = [] # Detailed rows from each employee sheet
    violations_list = []            # List of violation dictionaries

    # For start date validation, track the first instance per project per month.
    # Key: (project, month) where month is derived from "Status Date (Every Friday)"
    project_month_info = {}

    # Allowed values for each of the six last columns (exact, case sensitive)
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
    
    # For start date validation exceptions:
    # If either "Main project" or "Name of the Project" exactly equals one of these, skip the start date check.
    start_date_exceptions = ["Annual Leave"]

    # Process each employee's sheet
    for emp in employee_names:
        if emp not in all_sheet_names:
            st.warning(f"Sheet for employee '{emp}' not found. Skipping.")
            continue
        df = pd.read_excel(file_path, sheet_name=emp)
        # Normalize headers (remove newline characters, extra spaces)
        df.columns = [str(col).replace("\n", " ").strip() for col in df.columns]
        
        req_cols = ["Status Date (Every Friday)", "Main project", "Name of the Project", "Start Date", "Weekly Time Spent(Hrs)"]
        for col in req_cols:
            if col not in df.columns:
                st.error(f"Column '{col}' not found in sheet '{emp}'.")
                return None
        
        df["Employee"] = emp
        df["RowNumber"] = df.index + 2  # assuming header is in row 1

        # Convert "Status Date (Every Friday)" to datetime (format mm-dd-yyyy)
        df["Status Date (Every Friday)"] = pd.to_datetime(
            df["Status Date (Every Friday)"], format='%m-%d-%Y', errors='coerce'
        )
        
        # ----- Allowed Values Check for Last Six Columns -----
        for col, allowed in allowed_values.items():
            for i, val in df[col].iteritems():
                if pd.isna(val):
                    continue
                # Split on comma and strip tokens.
                tokens = [token.strip() for token in str(val).split(",") if token.strip()]
                if len(tokens) != 1 or tokens[0] not in allowed:
                    violations_list.append({
                        "Employee": emp,
                        "Violation Type": f"Invalid value in '{col}': found '{val}'",
                        "Location": f"Sheet '{emp}', Row {df.at[i, 'RowNumber']}"
                    })
        
        # ----- Start Date Validation (per project per month) -----
        for i, row in df.iterrows():
            project = row["Name of the Project"]
            start_date_val = row["Start Date"]
            main_project_val = str(row["Main project"]).strip() if not pd.isna(row["Main project"]) else ""
            project_val = str(project).strip() if not pd.isna(project) else ""
            if main_project_val in start_date_exceptions or project_val in start_date_exceptions:
                continue
            if pd.notna(project) and pd.notna(start_date_val) and pd.notna(row["Status Date (Every Friday)"]):
                month_key = row["Status Date (Every Friday)"].strftime("%Y-%m")
                key = (project, month_key)
                current_start = pd.to_datetime(start_date_val, format='%m-%d-%Y', errors='coerce')
                if key not in project_month_info:
                    project_month_info[key] = {"start_date": current_start, "sheet": emp, "row": row["RowNumber"]}
                else:
                    baseline = project_month_info[key]["start_date"]
                    # Only flag a violation if the current entry is not the first instance in that month.
                    if current_start != baseline:
                        violations_list.append({
                            "Employee": emp,
                            "Violation Type": (
                                f"Start date changed for project '{project}' in {month_key}: expected {baseline.strftime('%m-%d-%Y') if pd.notna(baseline) else 'NaT'}, "
                                f"found {current_start.strftime('%m-%d-%Y') if pd.notna(current_start) else 'NaT'} at Row {row['RowNumber']}"
                            ),
                            "Location": f"Sheet '{emp}', Row {row['RowNumber']}"
                        })
        
        # ----- Weekly Hours Validation -----
        df["Weekly Time Spent(Hrs)"] = pd.to_numeric(df["Weekly Time Spent(Hrs)"], errors='coerce').fillna(0)
        friday_rows = df[df["Status Date (Every Friday)"].dt.weekday == 4]
        unique_fridays = friday_rows["Status Date (Every Friday)"].dropna().unique()
        for friday in unique_fridays:
            week_start = friday - timedelta(days=4)
            week_df = df[(df["Status Date (Every Friday)"] >= week_start) &
                         (df["Status Date (Every Friday)"] <= friday)]
            week_hours = week_df["Weekly Time Spent(Hrs)"].sum()
            if week_hours < 40:
                violations_list.append({
                    "Employee": emp,
                    "Violation Type": (
                        f"Insufficient weekly work hours: {week_hours} (<40) for week ending {friday.strftime('%m-%d-%Y')} "
                        f"(from {week_start.strftime('%m-%d-%Y')} to {friday.strftime('%m-%d-%Y')})"
                    ),
                    "Location": f"Sheet '{emp}', Rows: {', '.join(map(str, week_df['RowNumber'].tolist()))}"
                })
        
        # ----- Monthly Summary for This Employee -----
        # For reporting, calculate separate totals for PTO and Work Hours.
        df["PTO Hours"] = df.apply(lambda r: r["Weekly Time Spent(Hrs)"] if "PTO" in str(r["Main project"]) else 0, axis=1)
        df["Work Hours"] = df.apply(lambda r: r["Weekly Time Spent(Hrs)"] if "PTO" not in str(r["Main project"]) else 0, axis=1)
        df["Month"] = df["Status Date (Every Friday)"].dt.to_period("M").astype(str)
        monthly_summary = df.groupby("Month").agg({"Work Hours": "sum", "PTO Hours": "sum"}).reset_index()
        monthly_summary["Employee"] = emp
        employee_reports_list.append(monthly_summary)
        
        working_hours_details_list.append(df)
    
    team_monthly_summary = pd.concat(employee_reports_list, ignore_index=True) if employee_reports_list else pd.DataFrame()
    working_hours_details = pd.concat(working_hours_details_list, ignore_index=True) if working_hours_details_list else pd.DataFrame()
    violations_df = pd.DataFrame(violations_list)
    
    # Create a full Excel workbook in memory (all sheets)
    full_output = BytesIO()
    with pd.ExcelWriter(full_output, engine="openpyxl") as writer:
        team_monthly_summary.to_excel(writer, sheet_name="Team_Monthly_Summary", index=False)
        working_hours_details.to_excel(writer, sheet_name="Working_Hours_Details", index=False)
        violations_df.to_excel(writer, sheet_name="Violations", index=False)
    full_output.seek(0)
    
    return team_monthly_summary, working_hours_details, violations_df, full_output

# ----------------- READ THE EXCEL FROM FIXED PATH -----------------
result = process_excel_file(FILE_PATH)
if result is None:
    st.error("Error processing file.")
else:
    team_monthly_summary, working_hours_details, violations_df, full_excel = result
    st.success("Reports generated successfully!")

    # ----------------- DASHBOARD TABS -----------------
    tabs = st.tabs(["Team Monthly Summary", "Working Hours Summary", "Violations"])

    # ----- Tab 1: Team Monthly Summary -----
    with tabs[0]:
        st.subheader("Team Monthly Summary")
        emp_filter = st.multiselect("Select Employee(s)",
                                    options=sorted(team_monthly_summary["Employee"].unique()),
                                    default=sorted(team_monthly_summary["Employee"].unique()))
        month_filter = st.multiselect("Select Month(s)",
                                      options=sorted(team_monthly_summary["Month"].unique()),
                                      default=sorted(team_monthly_summary["Month"].unique()))
        filtered_team = team_monthly_summary[
            (team_monthly_summary["Employee"].isin(emp_filter)) &
            (team_monthly_summary["Month"].isin(month_filter))
        ]
        st.dataframe(filtered_team, use_container_width=True)
        # Download filtered team monthly summary
        to_download_team = BytesIO()
        with pd.ExcelWriter(to_download_team, engine="openpyxl") as writer:
            filtered_team.to_excel(writer, sheet_name="Filtered_Team_Monthly", index=False)
        to_download_team.seek(0)
        st.download_button("Download Filtered Team Monthly Report",
                           data=to_download_team,
                           file_name="filtered_team_monthly.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ----- Tab 2: Working Hours Summary -----
    with tabs[1]:
        st.subheader("Working Hours Summary")
        emp_filter_wh = st.multiselect("Select Employee(s)",
                                       options=sorted(working_hours_details["Employee"].unique()),
                                       default=sorted(working_hours_details["Employee"].unique()))
        working_hours_details["Month"] = working_hours_details["Status Date (Every Friday)"].dt.to_period("M").astype(str)
        month_filter_wh = st.multiselect("Select Month(s)",
                                         options=sorted(working_hours_details["Month"].unique()),
                                         default=sorted(working_hours_details["Month"].unique()))
        week_filter = st.text_input("Enter Friday Date for Week Filter (mm-dd-yyyy)", value="")
        filtered_wh = working_hours_details[
            (working_hours_details["Employee"].isin(emp_filter_wh)) &
            (working_hours_details["Month"].isin(month_filter_wh))
        ]
        if week_filter:
            try:
                friday_date = datetime.strptime(week_filter, "%m-%d-%Y")
                filtered_wh = filtered_wh[filtered_wh["Status Date (Every Friday)"] == friday_date]
            except Exception:
                st.error("Invalid date format for week filter.")
        st.dataframe(filtered_wh, use_container_width=True)
        # Download filtered working hours summary
        to_download_wh = BytesIO()
        with pd.ExcelWriter(to_download_wh, engine="openpyxl") as writer:
            filtered_wh.to_excel(writer, sheet_name="Filtered_Working_Hours", index=False)
        to_download_wh.seek(0)
        st.download_button("Download Filtered Working Hours Report",
                           data=to_download_wh,
                           file_name="filtered_working_hours.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
    # ----- Tab 3: Violations -----
    with tabs[2]:
        st.subheader("Violations")
        if violations_df.empty:
            st.info("No violations found.")
        else:
            emp_filter_v = st.multiselect("Select Employee(s)",
                                          options=sorted(violations_df["Employee"].unique()),
                                          default=sorted(violations_df["Employee"].unique()))
            type_filter = st.multiselect("Select Violation Type(s)",
                                         options=sorted(violations_df["Violation Type"].unique()),
                                         default=sorted(violations_df["Violation Type"].unique()))
            filtered_v = violations_df[
                (violations_df["Employee"].isin(emp_filter_v)) &
                (violations_df["Violation Type"].isin(type_filter))
            ]
            st.dataframe(filtered_v, use_container_width=True)
            # Download filtered violations
            to_download_v = BytesIO()
            with pd.ExcelWriter(to_download_v, engine="openpyxl") as writer:
                filtered_v.to_excel(writer, sheet_name="Filtered_Violations", index=False)
            to_download_v.seek(0)
            st.download_button("Download Filtered Violations Report",
                               data=to_download_v,
                               file_name="filtered_violations.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
    # Download full Excel report (all sheets)
    st.markdown("---")
    st.subheader("Download Full Excel Report")
    st.download_button("Download Full Excel Report",
                       data=full_excel,
                       file_name="full_report.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")