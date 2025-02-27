import pandas as pd
from datetime import datetime
from openpyxl.styles import Alignment

# ----------------- CONFIGURATION -----------------
input_file = "input.xlsx"         # Change to your Excel file name
output_file = "output_validated.xlsx"

# Allowed values for the last six columns (exact and case sensitive)
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

# List of exception values (exact match) for start date validation.
# If either "Main project" or "Name of the Project" exactly equals one of these, skip the start date check.
start_date_exceptions = ["Annual Leave"]

# The first sheet named 'Home' contains the employee names in column F.
home_sheet = "Home"

# ----------------- READ THE EXCEL FILE -----------------
home_df = pd.read_excel(input_file, sheet_name=home_sheet, header=None)
employee_names = home_df.iloc[2:, 5].dropna().astype(str).tolist()

xls = pd.ExcelFile(input_file)
all_sheet_names = xls.sheet_names

# Containers for violations and monthly reports.
violations = []    # Each violation is a dict with Employee, Violation Type, and Location.
employee_reports = {}  # Key: employee name, Value: monthly summary DataFrame.
# For project start date tracking: map project name to its first encountered start date info.
project_start_info = {}

# ----------------- PROCESS EACH EMPLOYEE SHEET -----------------
for emp in employee_names:
    if emp not in all_sheet_names:
        print(f"Warning: No sheet found for employee '{emp}'. Skipping.")
        continue

    df = pd.read_excel(input_file, sheet_name=emp)
    # Normalize headers: replace newline characters with a space and strip extra spaces.
    df.columns = [str(col).replace("\n", " ").strip() for col in df.columns]

    required_columns = [
        "Status Date (Every Friday)", "Main project", "Name of the Project", "Start Date", 
        "Weekly Time Spent(Hrs)"
    ]
    for col in required_columns:
        if col not in df.columns:
            raise ValueError(f"Column '{col}' not found in sheet '{emp}'.")

    # Record original row numbers (assuming header is row 1)
    df["RowNumber"] = df.index + 2

    # Convert "Status Date (Every Friday)" using the known format (month-day-year)
    df["Status Date (Every Friday)"] = pd.to_datetime(
        df["Status Date (Every Friday)"], format='%m-%d-%Y', errors='coerce'
    )

    # Process each row for validations.
    for idx, row in df.iterrows():
        row_num = row["RowNumber"]

        # 1. Validate allowed values for each of the last six columns.
        for col, allowed in allowed_values.items():
            cell_val = row.get(col)
            if pd.notna(cell_val):
                # Split the cell value on commas, then strip each token.
                tokens = [token.strip() for token in str(cell_val).split(',') if token.strip()]
                # There must be exactly one token and it must be exactly in the allowed list.
                if len(tokens) != 1 or tokens[0] not in allowed:
                    violations.append({
                        "Employee": emp,
                        "Violation Type": f"Invalid value in '{col}': found '{cell_val}'",
                        "Location": f"Sheet '{emp}', Row {row_num}"
                    })

        # 3. Validate that the start date for a project is not changed.
        project_name = row.get("Name of the Project")
        start_date = row.get("Start Date")
        main_project_val = str(row.get("Main project")).strip() if pd.notna(row.get("Main project")) else ""
        project_name_val = str(project_name).strip() if pd.notna(project_name) else ""
        # Skip start date validation if either "Main project" or "Name of the Project" exactly matches an exception.
        if (main_project_val in start_date_exceptions) or (project_name_val in start_date_exceptions):
            continue

        if pd.notna(project_name) and pd.notna(start_date):
            start_date_converted = pd.to_datetime(start_date, format='%m-%d-%Y', errors='coerce')
            if project_name not in project_start_info:
                project_start_info[project_name] = {
                    "start_date": start_date_converted,
                    "sheet": emp,
                    "row": row_num
                }
            else:
                correct_info = project_start_info[project_name]
                if start_date_converted != correct_info["start_date"]:
                    expected_date_str = (correct_info["start_date"].strftime('%m-%d-%Y')
                                           if pd.notna(correct_info["start_date"]) else "NaT")
                    found_date_str = (start_date_converted.strftime('%m-%d-%Y')
                                      if pd.notna(start_date_converted) else "NaT")
                    violations.append({
                        "Employee": emp,
                        "Violation Type": (
                            f"Start date changed for project '{project_name}': expected {expected_date_str} "
                            f"(Sheet: {correct_info['sheet']}, Row: {correct_info['row']}) but found {found_date_str} at Row {row_num}"
                        ),
                        "Location": f"Sheet '{emp}', Row {row_num}"
                    })

    # 2. Validate weekly work hours.
    # For weekly totals, sum up the raw "Weekly Time Spent(Hrs)" (including PTO).
    df["Weekly Time Spent(Hrs)"] = pd.to_numeric(df["Weekly Time Spent(Hrs)"], errors='coerce').fillna(0)
    # For the weekly check, find rows where the status date is a Friday.
    friday_df = df[df["Status Date (Every Friday)"].dt.weekday == 4]
    unique_fridays = friday_df["Status Date (Every Friday)"].dropna().unique()

    for friday in unique_fridays:
        week_start = friday - pd.Timedelta(days=4)
        # Define the week as from week_start through Friday (inclusive).
        week_rows = df[(df["Status Date (Every Friday)"] >= week_start) & 
                       (df["Status Date (Every Friday)"] <= friday)]
        week_hours_sum = week_rows["Weekly Time Spent(Hrs)"].sum()
        if week_hours_sum < 40:
            affected_rows = ", ".join(str(x) for x in week_rows["RowNumber"].tolist())
            violation_message = (
                f"Insufficient weekly work hours: {week_hours_sum} (<40) for week ending {friday.strftime('%m-%d-%Y')} "
                f"(from {week_start.strftime('%m-%d-%Y')} to {friday.strftime('%m-%d-%Y')})"
            )
            violations.append({
                "Employee": emp,
                "Violation Type": violation_message,
                "Location": f"Sheet '{emp}', Rows: {affected_rows}"
            })

    # 4. Create monthly summary report for the employee.
    # For reporting, also calculate separate totals for non-PTO and PTO.
    df["PTO Hours"] = df.apply(lambda row: row["Weekly Time Spent(Hrs)"]
                               if "PTO" in str(row.get("Main project", "")) else 0, axis=1)
    df["Work Hours"] = df.apply(lambda row: row["Weekly Time Spent(Hrs)"]
                                if "PTO" not in str(row.get("Main project", "")) else 0, axis=1)
    df["Month"] = df["Status Date (Every Friday)"].dt.to_period("M")
    monthly_summary = df.groupby("Month").agg({
        "Work Hours": "sum",
        "PTO Hours": "sum"
    }).reset_index()
    monthly_summary["Month"] = monthly_summary["Month"].astype(str)
    employee_reports[emp] = monthly_summary

# ----------------- WRITE RESULTS TO NEW EXCEL FILE WITH LEFT ALIGNMENT -----------------
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    for emp, report_df in employee_reports.items():
        sheet_name = f"{emp}_Report"
        report_df.to_excel(writer, sheet_name=sheet_name, index=False)

    if violations:
        violations_df = pd.DataFrame(violations)
    else:
        violations_df = pd.DataFrame(columns=["Employee", "Violation Type", "Location"])
    violations_df.to_excel(writer, sheet_name="Violations", index=False)

    workbook = writer.book
    for ws in workbook.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                cell.alignment = Alignment(horizontal="left")

print(f"Validation complete. Output written to '{output_file}'.")