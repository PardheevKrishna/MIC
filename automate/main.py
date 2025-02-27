import pandas as pd
from datetime import datetime

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
# Here we assume that cell F2 is the header and employee names start at row 3.
home_df = pd.read_excel(input_file, sheet_name=home_sheet, header=None)
employee_names = home_df.iloc[2:, 5].dropna().astype(str).tolist()

# Get all sheet names from the workbook
xls = pd.ExcelFile(input_file)
all_sheet_names = xls.sheet_names

# Containers for violations and employee monthly reports.
violations = []  # Each violation is a dict with Employee, Violation Type, and Location.
employee_reports = {}  # Key: employee name, Value: monthly summary DataFrame.
project_start_dates = {}  # Global dictionary to track the first start date for each project

# ----------------- PROCESS EACH EMPLOYEE SHEET -----------------
for emp in employee_names:
    if emp not in all_sheet_names:
        print(f"Warning: No sheet found for employee '{emp}'. Skipping.")
        continue

    df = pd.read_excel(input_file, sheet_name=emp)
    # Clean up column headers (strip extra spaces)
    df.columns = [str(col).strip() for col in df.columns]

    # Ensure required columns are present
    required_columns = [
        "Status Date (Every Friday)", "Main project", "Name of the Project", "Start Date", 
        "Weekly Time Spent(Hrs)"
    ]
    for col in required_columns:
        if col not in df.columns:
            raise ValueError(f"Column '{col}' not found in sheet '{emp}'.")

    # Process each row in the employee's sheet
    for idx, row in df.iterrows():
        # Excel row number (assumes header is row 1)
        row_num = idx + 2

        # ---- 1. Validate allowed values for the last six columns ----
        for col, allowed in allowed_values.items():
            cell_val = row.get(col)
            # Only flag if the cell is non-empty and does not exactly match one of the allowed values.
            if pd.notna(cell_val) and cell_val not in allowed:
                violations.append({
                    "Employee": emp,
                    "Violation Type": f"Invalid value in '{col}': found '{cell_val}'",
                    "Location": f"Sheet '{emp}', Row {row_num}"
                })

        # ---- 2. Validate weekly hours (each week must have at least 40 hours) ----
        try:
            weekly_hours = float(row.get("Weekly Time Spent(Hrs)", 0))
        except Exception:
            weekly_hours = 0
        if weekly_hours < 40:
            violations.append({
                "Employee": emp,
                "Violation Type": f"Insufficient weekly hours: {weekly_hours} (< 40)",
                "Location": f"Sheet '{emp}', Row {row_num}"
            })

        # ---- 3. Validate that the start date for a project is not changed ----
        project_name = row.get("Name of the Project")
        start_date = row.get("Start Date")
        if pd.notna(project_name) and pd.notna(start_date):
            # Convert start_date to a datetime object if it is not already
            if not isinstance(start_date, (datetime, pd.Timestamp)):
                try:
                    start_date = pd.to_datetime(start_date)
                except Exception:
                    pass
            # Check against global project_start_dates
            if project_name not in project_start_dates:
                project_start_dates[project_name] = start_date
            else:
                if start_date != project_start_dates[project_name]:
                    violations.append({
                        "Employee": emp,
                        "Violation Type": f"Start date changed for project '{project_name}' (expected {project_start_dates[project_name]}, got {start_date})",
                        "Location": f"Sheet '{emp}', Row {row_num}"
                    })

    # ---- 4. Create monthly summary report for the employee ----
    # Convert "Status Date (Every Friday)" to datetime; if conversion fails, errors become NaT.
    df["Status Date (Every Friday)"] = pd.to_datetime(df["Status Date (Every Friday)"], errors="coerce")
    # Create a "Month" column (e.g., "2025-02")
    df["Month"] = df["Status Date (Every Friday)"].dt.to_period("M")
    
    # Determine which rows are PTO: if "Main project" contains 'PTO'
    df["PTO Hours"] = df.apply(lambda row: row["Weekly Time Spent(Hrs)"]
                               if "PTO" in str(row.get("Main project", "")) else 0, axis=1)
    # Work hours are those rows where it is not PTO.
    df["Work Hours"] = df.apply(lambda row: 0 if "PTO" in str(row.get("Main project", "")) 
                                else row["Weekly Time Spent(Hrs)"], axis=1)
    
    # Group by Month to sum up Work and PTO hours
    monthly_summary = df.groupby("Month").agg({
        "Work Hours": "sum",
        "PTO Hours": "sum"
    }).reset_index()
    # Convert month to string format for reporting
    monthly_summary["Month"] = monthly_summary["Month"].astype(str)
    employee_reports[emp] = monthly_summary

# ----------------- WRITE RESULTS TO NEW EXCEL FILE -----------------
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    # Write one sheet per employee containing their monthly report.
    for emp, report_df in employee_reports.items():
        sheet_name = f"{emp}_Report"
        # Ensure sheet name fits Excel limitations.
        report_df.to_excel(writer, sheet_name=sheet_name, index=False)

    # Write all violations to a new "Violations" sheet.
    if violations:
        violations_df = pd.DataFrame(violations)
    else:
        violations_df = pd.DataFrame(columns=["Employee", "Violation Type", "Location"])
    violations_df.to_excel(writer, sheet_name="Violations", index=False)

print(f"Validation complete. Output written to '{output_file}'.")