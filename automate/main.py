import pandas as pd
from datetime import datetime
from openpyxl.styles import Alignment

# ----------------- CONFIGURATION -----------------
# Input and output file names
input_file = "input.xlsx"         # Change this to your Excel file name
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

# The first sheet named 'Home' contains the employee names in column F.
home_sheet = "Home"

# ----------------- READ THE EXCEL FILE -----------------
# Load the Home sheet to get employee names.
# We assume that cell F2 is the header and employee names start at row 3.
home_df = pd.read_excel(input_file, sheet_name=home_sheet, header=None)
employee_names = home_df.iloc[2:, 5].dropna().astype(str).tolist()

# Get all sheet names from the workbook
xls = pd.ExcelFile(input_file)
all_sheet_names = xls.sheet_names

# Containers for violations and employee monthly reports.
violations = []  # Each violation is a dict with Employee, Violation Type, and Location.
employee_reports = {}  # Key: employee name, Value: monthly summary DataFrame.
# For project start date tracking: store a dict mapping project name to its first encountered
# start date along with sheet and row number.
project_start_info = {}

# ----------------- PROCESS EACH EMPLOYEE SHEET -----------------
for emp in employee_names:
    if emp not in all_sheet_names:
        print(f"Warning: No sheet found for employee '{emp}'. Skipping.")
        continue

    df = pd.read_excel(input_file, sheet_name=emp)
    # Normalize column headers: replace newline characters with a space and strip extra spaces.
    df.columns = [str(col).replace("\n", " ").strip() for col in df.columns]

    # Ensure required columns are present.
    required_columns = [
        "Status Date (Every Friday)", "Main project", "Name of the Project", "Start Date", 
        "Weekly Time Spent(Hrs)"
    ]
    for col in required_columns:
        if col not in df.columns:
            raise ValueError(f"Column '{col}' not found in sheet '{emp}'.")

    # Add a column for original row numbers (assuming header is row 1)
    df["RowNumber"] = df.index + 2

    # ---- Convert dates using known format (month-day-year) ----
    # For "Status Date (Every Friday)"
    df["Status Date (Every Friday)"] = pd.to_datetime(
        df["Status Date (Every Friday)"], format='%m-%d-%Y', errors='coerce'
    )
    
    # Process each row for validations
    for idx, row in df.iterrows():
        row_num = row["RowNumber"]

        # ---- 1. Validate allowed values for the last six columns ----
        for col, allowed in allowed_values.items():
            cell_val = row.get(col)
            if pd.notna(cell_val) and cell_val not in allowed:
                violations.append({
                    "Employee": emp,
                    "Violation Type": f"Invalid value in '{col}': found '{cell_val}'",
                    "Location": f"Sheet '{emp}', Row {row_num}"
                })

        # ---- 3. Validate that the start date for a project is not changed ----
        project_name = row.get("Name of the Project")
        start_date = row.get("Start Date")
        if pd.notna(project_name) and pd.notna(start_date):
            # Convert start_date to datetime using the known format
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
                    violations.append({
                        "Employee": emp,
                        "Violation Type": (
                            f"Start date changed for project '{project_name}': expected {correct_info['start_date'].strftime('%m-%d-%Y')} "
                            f"(Sheet: {correct_info['sheet']}, Row: {correct_info['row']}) but found {start_date_converted.strftime('%m-%d-%Y')} at Row {row_num}"
                        ),
                        "Location": f"Sheet '{emp}', Row {row_num}"
                    })

    # ---- 2. Validate weekly work hours (each week must have at least 40 work hours) ----
    # First, compute Work Hours per row: if "PTO" is in "Main project", then work hours are 0.
    df["Work Hours"] = df.apply(
        lambda row: 0 if "PTO" in str(row.get("Main project", "")) else row["Weekly Time Spent(Hrs)"],
        axis=1
    )
    # Also compute PTO Hours for reporting (though these don't count toward 40 hours)
    df["PTO Hours"] = df.apply(
        lambda row: row["Weekly Time Spent(Hrs)"] if "PTO" in str(row.get("Main project", "")) else 0,
        axis=1
    )
    # Group rows by the Status Date (i.e. each week; the date is the Friday for that week)
    weekly_groups = df.groupby("Status Date (Every Friday)")
    for week_date, group in weekly_groups:
        # Sum up the work hours for the week (exclude PTO hours)
        week_work_sum = group["Work Hours"].sum()
        if week_work_sum < 40:
            # List affected rows (row numbers) in this group
            affected_rows = ", ".join(str(x) for x in group["RowNumber"].tolist())
            violations.append({
                "Employee": emp,
                "Violation Type": f"Insufficient weekly work hours: {week_work_sum} (<40) for week ending {week_date.strftime('%m-%d-%Y')}",
                "Location": f"Sheet '{emp}', Rows: {affected_rows}"
            })

    # ---- 4. Create monthly summary report for the employee ----
    # Create a "Month" column (e.g., "2025-02")
    df["Month"] = df["Status Date (Every Friday)"].dt.to_period("M")
    # Group by Month to sum up Work and PTO hours
    monthly_summary = df.groupby("Month").agg({
        "Work Hours": "sum",
        "PTO Hours": "sum"
    }).reset_index()
    # Convert month to string for reporting
    monthly_summary["Month"] = monthly_summary["Month"].astype(str)
    employee_reports[emp] = monthly_summary

# ----------------- WRITE RESULTS TO NEW EXCEL FILE WITH LEFT ALIGNMENT -----------------
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    # Write one sheet per employee containing their monthly report.
    for emp, report_df in employee_reports.items():
        sheet_name = f"{emp}_Report"
        report_df.to_excel(writer, sheet_name=sheet_name, index=False)

    # Write all violations to a new "Violations" sheet.
    if violations:
        violations_df = pd.DataFrame(violations)
    else:
        violations_df = pd.DataFrame(columns=["Employee", "Violation Type", "Location"])
    violations_df.to_excel(writer, sheet_name="Violations", index=False)

    # After writing all sheets, iterate over every cell in each worksheet to set left alignment.
    workbook = writer.book
    for ws in workbook.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                cell.alignment = Alignment(horizontal="left")
    # The workbook is saved when the context manager exits.

print(f"Validation complete. Output written to '{output_file}'.")