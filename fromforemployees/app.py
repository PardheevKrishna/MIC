from flask import Flask, render_template, request
from datetime import datetime, timedelta
import openpyxl
import os
import logging
from openpyxl.styles import PatternFill

app = Flask(__name__)
logging.basicConfig(level=logging.INFO)

LOGS_FILE = 'logs.xlsx'
LEAVES_FILE = 'leaves.xlsx'

# Orange fill to highlight anomaly rows
ORANGE_FILL = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")

def validate_fields(data):
    """
    Validate that each form field entry is valid.
    Return a list of error messages. (These errors block submission.)
    """
    errors = []

    # Validate Status Date (must be a valid date and a Friday)
    try:
        status_date = datetime.strptime(data['status_date'], "%Y-%m-%d").date()
        if status_date.weekday() != 4:
            errors.append("Status Date must be a Friday.")
    except ValueError:
        errors.append("Invalid Status Date format.")

    # Required text fields (Main Project, Project Name, Project Key Milestones, TM)
    if not data['main_project'].strip():
        errors.append("Main Project is required.")
    if not data['project_name'].strip():
        errors.append("Name of the Project is required.")
    if not data['project_key_milestones'].strip():
        errors.append("Project Key Milestones is required.")
    if not data['tm'].strip():
        errors.append("Team Member (TM) is required.")

    # Validate Start Date and Completion Date and that start is not after completion
    try:
        start_date = datetime.strptime(data['start_date'], "%Y-%m-%d").date()
    except ValueError:
        errors.append("Invalid Start Date format.")
    try:
        completion_date = datetime.strptime(data['completion_date'], "%Y-%m-%d").date()
    except ValueError:
        errors.append("Invalid Completion Date format.")
    else:
        try:
            start_date = datetime.strptime(data['start_date'], "%Y-%m-%d").date()
            if start_date > completion_date:
                errors.append("Start Date cannot be after Completion Date.")
        except:
            pass

    # Validate % of Completion (numeric, between 0 and 100)
    try:
        pct = float(data['percent_completion'])
        if pct < 0 or pct > 100:
            errors.append("% of Completion must be between 0 and 100.")
    except ValueError:
        errors.append("Invalid % of Completion.")

    # Validate Status (non-empty)
    if not data['status'].strip():
        errors.append("Status is required.")

    # Validate Weekly Time Spent (numeric, non-negative)
    try:
        wts = float(data['weekly_time_spent'])
        if wts < 0:
            errors.append("Weekly Time Spent cannot be negative.")
    except ValueError:
        errors.append("Invalid Weekly Time Spent value.")

    # Validate Projected Hours (numeric, non-negative)
    try:
        ph = float(data['projected_hours'])
        if ph < 0:
            errors.append("Projected Hours cannot be negative.")
    except ValueError:
        errors.append("Invalid Projected Hours value.")

    # Validate Functional Area, Project Category, Output Type, Impact Type (non-empty)
    if not data['functional_area'].strip():
        errors.append("Functional Area is required.")
    if not data['project_category'].strip():
        errors.append("Project Category is required.")
    if not data['output_type'].strip():
        errors.append("Output Type is required.")
    if not data['impact_type'].strip():
        errors.append("Impact Type is required.")

    # Validate Complexity (must be H, M, or L)
    if data['complexity'] not in ["H", "M", "L"]:
        errors.append("Complexity must be one of H, M, or L.")

    # Validate Novelty (must be one of the allowed options)
    if data['novelty'] not in ["BAU repetitive", "One time repetitive", "New one time"]:
        errors.append("Novelty must be one of: BAU repetitive, One time repetitive, or New one time.")

    return errors

def count_leave_days(tm, week_date):
    """
    Count how many leave days the given Team Member (TM) has in the week (Monday to Friday) of week_date.
    """
    leave_count = 0
    if os.path.exists(LEAVES_FILE):
        try:
            wb = openpyxl.load_workbook(LEAVES_FILE)
            sheet = wb.active
            monday = week_date - timedelta(days=week_date.weekday())
            friday = monday + timedelta(days=4)
            for row in sheet.iter_rows(min_row=2, values_only=True):
                # Expecting leaves.xlsx columns: [TM, Date, Leave Type]
                if str(row[0]).strip() == tm:
                    try:
                        leave_date = datetime.strptime(str(row[1]).strip(), "%Y-%m-%d").date()
                        if monday <= leave_date <= friday:
                            leave_count += 1
                    except:
                        continue
        except Exception as e:
            logging.error("Error reading leaves.xlsx: %s", e)
    return leave_count

def check_anomalies(data):
    """
    Check for fake data only related to working hours.
    Specifically, if the Weekly Time Spent (Hrs) exceeds the maximum allowed
    based on the number of working days available (i.e. 9 hours per day for Monday–Friday,
    minus any leave days for that week).
    For example, if a TM is on leave for 2 days, the maximum allowed is 9 * (5-2) = 27 hrs.
    """
    anomalies = []
    try:
        status_date = datetime.strptime(data['status_date'], "%Y-%m-%d").date()
        weekly_time_spent = float(data['weekly_time_spent'])
    except:
        return anomalies  # If parsing fails, skip anomaly check

    tm = data['tm'].strip()
    leave_days = count_leave_days(tm, status_date)
    allowed_days = 5 - leave_days
    max_allowed_hours = allowed_days * 9

    if weekly_time_spent > max_allowed_hours:
        anomalies.append(f"Weekly Time Spent exceeds allowed limit. With {leave_days} leave day(s), maximum allowed is {max_allowed_hours} hrs.")
    return anomalies

def append_log(data, anomaly_message):
    """
    Append a new row into logs.xlsx with the 17 fields and the Anomaly Reason.
    If an anomaly exists, highlight the entire row in orange.
    """
    try:
        if os.path.exists(LOGS_FILE):
            wb = openpyxl.load_workbook(LOGS_FILE)
            sheet = wb.active
        else:
            wb = openpyxl.Workbook()
            sheet = wb.active
            headers = [
                "Status Date (Fri)",
                "Main Project",
                "Project Name",
                "Project Key Milestones",
                "Team Member (TM)",
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
            ]
            sheet.append(headers)

        new_row = sheet.max_row + 1
        sheet.cell(row=new_row, column=1, value=data['status_date'])
        sheet.cell(row=new_row, column=2, value=data['main_project'])
        sheet.cell(row=new_row, column=3, value=data['project_name'])
        sheet.cell(row=new_row, column=4, value=data['project_key_milestones'])
        sheet.cell(row=new_row, column=5, value=data['tm'])
        sheet.cell(row=new_row, column=6, value=data['start_date'])
        sheet.cell(row=new_row, column=7, value=data['completion_date'])
        
        cell_pct = sheet.cell(row=new_row, column=8)
        cell_pct.value = float(data['percent_completion'])
        cell_pct.number_format = '0.00'
        
        sheet.cell(row=new_row, column=9, value=data['status'])
        
        cell_wts = sheet.cell(row=new_row, column=10)
        cell_wts.value = float(data['weekly_time_spent'])
        cell_wts.number_format = '0.00'
        
        cell_ph = sheet.cell(row=new_row, column=11)
        cell_ph.value = float(data['projected_hours'])
        cell_ph.number_format = '0.00'
        
        sheet.cell(row=new_row, column=12, value=data['functional_area'])
        sheet.cell(row=new_row, column=13, value=data['project_category'])
        sheet.cell(row=new_row, column=14, value=data['complexity'])
        sheet.cell(row=new_row, column=15, value=data['novelty'])
        sheet.cell(row=new_row, column=16, value=data['output_type'])
        sheet.cell(row=new_row, column=17, value=data['impact_type'])
        sheet.cell(row=new_row, column=18, value=anomaly_message)
        
        # If any anomaly exists, highlight the row in orange.
        if anomaly_message:
            for col in range(1, 19):
                sheet.cell(row=new_row, column=col).fill = ORANGE_FILL

        wb.save(LOGS_FILE)
        logging.info("Log appended successfully.")
        return True, None
    except Exception as e:
        logging.error("Error appending log: %s", e)
        return False, str(e)

@app.route("/", methods=["GET"])
def index():
    return render_template("form.html")

@app.route("/submit", methods=["POST"])
def submit():
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

    # First, validate that the field entries are valid.
    errors = validate_fields(data)
    if errors:
        return render_template("form.html", error="; ".join(errors))

    # Then, check for anomalies only related to fake data (e.g. working hours limits)
    anomalies = check_anomalies(data)
    anomaly_message = "; ".join(anomalies) if anomalies else ""

    success, err_msg = append_log(data, anomaly_message)
    if success:
        return render_template("form.html", message="Log submitted successfully!")
    else:
        return render_template("form.html", error=f"Error saving log: {err_msg}")

if __name__ == "__main__":
    app.run(debug=True)