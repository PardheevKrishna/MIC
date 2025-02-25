from flask import Flask, render_template, request
from datetime import datetime
import openpyxl
import os
import logging

from openpyxl.styles import PatternFill

app = Flask(__name__)
logging.basicConfig(level=logging.INFO)

LOGS_FILE = 'logs.xlsx'

# For highlighting anomalies
FILL_ORANGE = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")

# Enumerations for validation
VALID_FUNCTIONAL_AREAS = [
    "CRIT",
    "CRIT - Data Management",
    "CRIT - Data Governance",
    "CRIT - Regulatory Reporting",
    "CRIT - Portfolio Reporting",
    "CRIT - Transformation"
]

VALID_PROJECT_CATEGORIES = [
    "Data Infrastructure",
    "Monitoring & Insights",
    "Analytics / Strategy Development",
    "GDA Related",
    "Trainings and Team Meeting"
]

VALID_COMPLEXITY = ["H", "M", "L"]

VALID_NOVELTY = [
    "BAU repetitive",
    "One time repetitive",
    "New one time"
]

VALID_OUTPUT_TYPES = [
    "Core production work",
    "Ad-hoc long-term projects",
    "Ad-hoc short-term projects",
    "Business Management",
    "Administration",
    "Trainings/L&D activities",
    "Others"
]

VALID_IMPACT_TYPES = [
    "Customer Experience",
    "Financial impact",
    "Insights",
    "Risk reduction",
    "Others"
]

def check_data_integrity(data):
    """
    Validate fields and return a list of anomaly messages.
    We do not block submission; anomalies go into logs.
    """
    anomalies = []

    # 1) Status Date must be a Friday
    try:
        status_dt = datetime.strptime(data['status_date'], "%Y-%m-%d").date()
        if status_dt.weekday() != 4:  # Monday=0 ... Sunday=6
            anomalies.append("Status Date is not a Friday.")
    except ValueError:
        anomalies.append("Invalid Status Date format.")

    # 2) Main project -> required, check length
    if not data['main_project'].strip():
        anomalies.append("Main Project is empty.")

    # 3) Name of the Project -> required
    if not data['project_name'].strip():
        anomalies.append("Name of the Project is empty.")

    # 4) Project Key Milestones -> required, check length
    if not data['project_key_milestones'].strip():
        anomalies.append("Project Key Milestones is empty.")

    # 5) TM -> required
    if not data['tm'].strip():
        anomalies.append("TM field is empty.")

    # 6 & 7) Start Date <= Completion Date
    try:
        start_dt = datetime.strptime(data['start_date'], "%Y-%m-%d").date()
        completion_dt = datetime.strptime(data['completion_date'], "%Y-%m-%d").date()
        if start_dt > completion_dt:
            anomalies.append("Start Date is after Completion Date.")
    except ValueError:
        anomalies.append("Invalid Start or Completion Date format.")

    # 8) % of Completion -> 0 <= x <= 100
    try:
        pc = float(data['percent_completion'])
        if pc < 0 or pc > 100:
            anomalies.append("% of Completion must be between 0 and 100.")
    except ValueError:
        anomalies.append("Invalid % of Completion.")

    # 9) Status -> required
    if not data['status'].strip():
        anomalies.append("Status is empty.")

    # 10) Weekly Time Spent(Hrs) -> numeric, non-negative
    try:
        wts = float(data['weekly_time_spent'])
        if wts < 0:
            anomalies.append("Weekly Time Spent cannot be negative.")
    except ValueError:
        anomalies.append("Invalid Weekly Time Spent.")

    # 11) Projected hours -> numeric, non-negative
    try:
        ph = float(data['projected_hours'])
        if ph < 0:
            anomalies.append("Projected Hours cannot be negative.")
        # Optional: If weekly time spent > projected hours => anomaly
        if 'wts' in locals() and wts > ph:
            anomalies.append("Weekly Time Spent exceeds Projected Hours.")
    except ValueError:
        anomalies.append("Invalid Projected Hours.")

    # 12) Functional Area -> must be in valid list
    if data['functional_area'] not in VALID_FUNCTIONAL_AREAS:
        anomalies.append("Invalid Functional Area selected.")

    # 13) Project Category -> must be in valid list
    if data['project_category'] not in VALID_PROJECT_CATEGORIES:
        anomalies.append("Invalid Project Category selected.")

    # 14) Complexity -> must be H, M, or L
    if data['complexity'] not in VALID_COMPLEXITY:
        anomalies.append("Invalid Complexity value.")

    # 15) Novelty -> must be in valid list
    if data['novelty'] not in VALID_NOVELTY:
        anomalies.append("Invalid Novelty value.")

    # 16) Output Type -> must be in valid list
    if data['output_type'] not in VALID_OUTPUT_TYPES:
        anomalies.append("Invalid Output Type selected.")

    # 17) Impact Type -> must be in valid list
    if data['impact_type'] not in VALID_IMPACT_TYPES:
        anomalies.append("Invalid Impact Type selected.")

    return anomalies

def append_log(data, anomaly_message):
    """
    Append a new row into logs.xlsx, including an anomaly reason if any.
    If an anomaly exists, we can highlight the row in orange.
    """
    try:
        if os.path.exists(LOGS_FILE):
            wb = openpyxl.load_workbook(LOGS_FILE)
            sheet = wb.active
        else:
            wb = openpyxl.Workbook()
            sheet = wb.active
            sheet.append([
                "Status Date (Fri)",
                "Main Project",
                "Project Name",
                "Project Key Milestones",
                "TM",
                "Start Date",
                "Completion Date",
                "% of Completion",
                "Status",
                "Weekly Time Spent(Hrs)",
                "Projected Hours",
                "Functional Area",
                "Project Category",
                "Complexity",
                "Novelty",
                "Output Type",
                "Impact Type",
                "Anomaly Reason"
            ])

        new_row = sheet.max_row + 1
        # Insert each field
        sheet.cell(row=new_row, column=1,  value=data['status_date'])
        sheet.cell(row=new_row, column=2,  value=data['main_project'])
        sheet.cell(row=new_row, column=3,  value=data['project_name'])
        sheet.cell(row=new_row, column=4,  value=data['project_key_milestones'])
        sheet.cell(row=new_row, column=5,  value=data['tm'])
        sheet.cell(row=new_row, column=6,  value=data['start_date'])
        sheet.cell(row=new_row, column=7,  value=data['completion_date'])

        # % completion as numeric
        c_percent = sheet.cell(row=new_row, column=8)
        c_percent.value = float(data['percent_completion'])
        c_percent.number_format = '0.00'

        sheet.cell(row=new_row, column=9,  value=data['status'])

        # Weekly Time Spent as numeric
        c_wts = sheet.cell(row=new_row, column=10)
        c_wts.value = float(data['weekly_time_spent'])
        c_wts.number_format = '0.00'

        # Projected Hours as numeric
        c_ph = sheet.cell(row=new_row, column=11)
        c_ph.value = float(data['projected_hours'])
        c_ph.number_format = '0.00'

        sheet.cell(row=new_row, column=12, value=data['functional_area'])
        sheet.cell(row=new_row, column=13, value=data['project_category'])
        sheet.cell(row=new_row, column=14, value=data['complexity'])
        sheet.cell(row=new_row, column=15, value=data['novelty'])
        sheet.cell(row=new_row, column=16, value=data['output_type'])
        sheet.cell(row=new_row, column=17, value=data['impact_type'])
        sheet.cell(row=new_row, column=18, value=anomaly_message)

        # Highlight row if anomalies exist
        if anomaly_message:
            for col_idx in range(1, 19):
                sheet.cell(row=new_row, column=col_idx).fill = FILL_ORANGE

        wb.save(LOGS_FILE)
        logging.info("Data appended successfully.")
        return True, None
    except Exception as e:
        logging.error("Error appending log: %s", e)
        return False, str(e)

@app.route("/", methods=["GET"])
def index():
    return render_template("form.html")

@app.route("/submit", methods=["POST"])
def submit():
    try:
        data = {
            "status_date": request.form.get("status_date", ""),
            "main_project": request.form.get("main_project", ""),
            "project_name": request.form.get("project_name", ""),
            "project_key_milestones": request.form.get("project_key_milestones", ""),
            "tm": request.form.get("tm", ""),
            "start_date": request.form.get("start_date", ""),
            "completion_date": request.form.get("completion_date", ""),
            "percent_completion": request.form.get("percent_completion", ""),
            "status": request.form.get("status", ""),
            "weekly_time_spent": request.form.get("weekly_time_spent", ""),
            "projected_hours": request.form.get("projected_hours", ""),
            "functional_area": request.form.get("functional_area", ""),
            "project_category": request.form.get("project_category", ""),
            "complexity": request.form.get("complexity", ""),
            "novelty": request.form.get("novelty", ""),
            "output_type": request.form.get("output_type", ""),
            "impact_type": request.form.get("impact_type", "")
        }

        anomalies = check_data_integrity(data)
        anomaly_message = "; ".join(anomalies) if anomalies else ""

        success, err_msg = append_log(data, anomaly_message)
        if success:
            # Show only a generic success message, no matter if anomalies exist
            return render_template("form.html", message="Log submitted successfully!")
        else:
            return render_template("form.html", error=f"Error saving log: {err_msg}")
    except Exception as e:
        logging.error("Error in submit route: %s", e)
        return render_template("form.html", error="An unexpected error occurred.")

if __name__ == "__main__":
    app.run(debug=True)