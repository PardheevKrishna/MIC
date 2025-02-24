from flask import Flask, render_template, request
from datetime import datetime, timedelta
import openpyxl
import os
import logging

# Import PatternFill for cell background coloring
from openpyxl.styles import PatternFill

app = Flask(__name__)
logging.basicConfig(level=logging.INFO)

# File paths
LOGS_FILE = 'logs.xlsx'
LEAVES_FILE = 'leaves.xlsx'

# Example public holidays for demonstration
PUBLIC_HOLIDAYS = ["2025-02-14"]

def load_projects():
    """Load unique project names from logs.xlsx."""
    projects = set()
    if os.path.exists(LOGS_FILE):
        try:
            wb = openpyxl.load_workbook(LOGS_FILE)
            sheet = wb.active
            for row in sheet.iter_rows(min_row=2, values_only=True):
                # row[3] is the 'Project' column
                if row[3]:
                    projects.add(row[3])
        except Exception as e:
            logging.error("Error reading logs.xlsx: %s", e)
    return sorted(projects)

def duplicate_entry(data):
    """Check if a log for the same employee on the same date already exists."""
    if os.path.exists(LOGS_FILE):
        try:
            wb = openpyxl.load_workbook(LOGS_FILE)
            sheet = wb.active
            for row in sheet.iter_rows(min_row=2, values_only=True):
                # row[0] = Employee ID, row[2] = Date
                if row[0] == data['employee_id'] and str(row[2]) == data['date']:
                    return True
        except Exception as e:
            logging.error("Error checking duplicate entry: %s", e)
    return False

def get_weekly_hours(employee_id, date_str):
    """Sum the hours logged in the week of the given date for the specified employee."""
    total = 0.0
    try:
        new_date = datetime.strptime(date_str, "%Y-%m-%d").date()
    except Exception:
        return total
    
    # Calculate Monday and Sunday of that week
    monday = new_date - timedelta(days=new_date.weekday())
    sunday = monday + timedelta(days=6)
    
    if os.path.exists(LOGS_FILE):
        try:
            wb = openpyxl.load_workbook(LOGS_FILE)
            sheet = wb.active
            for row in sheet.iter_rows(min_row=2, values_only=True):
                if row[0] == employee_id:
                    try:
                        row_date = (
                            datetime.strptime(str(row[2]), "%Y-%m-%d").date()
                            if isinstance(row[2], str) else row[2]
                        )
                    except Exception:
                        continue
                    if monday <= row_date <= sunday:
                        try:
                            total += float(row[4])
                        except Exception:
                            continue
        except Exception as e:
            logging.error("Error reading weekly hours: %s", e)
    return total

def check_leave(data):
    """
    Check if the employee is on leave on the given date (from leaves.xlsx).
    Return an anomaly message if found.
    """
    if os.path.exists(LEAVES_FILE):
        try:
            wb = openpyxl.load_workbook(LEAVES_FILE)
            sheet = wb.active
            for row in sheet.iter_rows(min_row=2, values_only=True):
                # columns: Employee ID, Date, Leave Type
                if str(row[0]).strip() == data['employee_id'] and str(row[1]).strip() == data['date']:
                    leave_type = row[2] if row[2] else "Leave"
                    return f"Employee is marked as on leave ({leave_type}) on this date."
        except Exception as e:
            logging.error("Error reading leaves.xlsx: %s", e)
    return ""

def check_data_integrity(data):
    """
    Validate the submitted log data and return a list of anomaly messages.
    These do not block submission; they're recorded in the logs.
    """
    anomalies = []
    
    # Check hours is numeric
    try:
        hours = float(data['hours'])
    except ValueError:
        anomalies.append("Invalid number for hours worked.")
        return anomalies  # Can't parse hours, so just return here
    
    # Validate date
    try:
        date_obj = datetime.strptime(data['date'], "%Y-%m-%d").date()
    except Exception:
        anomalies.append("Invalid date format.")
        return anomalies
    
    # Working Day Check
    if date_obj.weekday() >= 5:
        anomalies.append("Logged date falls on a weekend.")
    if data['date'] in PUBLIC_HOLIDAYS:
        anomalies.append("Logged date is a public holiday.")

    # Duplicate Entry
    if duplicate_entry(data):
        anomalies.append("Duplicate log entry found for this employee on the given date.")

    # Daily Hours Check: limit of 9 hours
    if hours > 9:
        anomalies.append("Logged hours exceed the typical daily limit of 9 hours.")

    # Weekly Hours Check: limit of 40 hours
    weekly_hours = get_weekly_hours(data['employee_id'], data['date'])
    if (weekly_hours + hours) > 40:
        anomalies.append("Total weekly hours exceed the corporate limit of 40 hours.")

    # Leave Check
    leave_anomaly = check_leave(data)
    if leave_anomaly:
        anomalies.append(leave_anomaly)

    return anomalies

def append_log(data, anomaly_message):
    """
    Append a new row into logs.xlsx, including an anomaly reason if any.
    Also ensure hours are stored as numeric and aligned properly in Excel.
    If an anomaly exists, highlight the entire row in orange.
    """
    try:
        if os.path.exists(LOGS_FILE):
            wb = openpyxl.load_workbook(LOGS_FILE)
            sheet = wb.active
        else:
            wb = openpyxl.Workbook()
            sheet = wb.active
            sheet.append(["Employee ID", "Employee Name", "Date", "Project", "Hours Worked", "Description", "Anomaly Reason"])
        
        new_row = sheet.max_row + 1
        
        # Fill cell values
        sheet.cell(row=new_row, column=1, value=data['employee_id'])
        sheet.cell(row=new_row, column=2, value=data['employee_name'])
        sheet.cell(row=new_row, column=3, value=data['date'])
        sheet.cell(row=new_row, column=4, value=data['project'])
        
        # Convert hours to float and set number format to ensure numeric alignment
        hours_cell = sheet.cell(row=new_row, column=5)
        hours_cell.value = float(data['hours'])
        hours_cell.number_format = '0.00'
        
        sheet.cell(row=new_row, column=6, value=data['description'])
        sheet.cell(row=new_row, column=7, value=anomaly_message)
        
        # If there's an anomaly, highlight the entire row in orange
        if anomaly_message:
            fill_orange = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
            for col in range(1, 8):
                sheet.cell(row=new_row, column=col).fill = fill_orange
        
        wb.save(LOGS_FILE)
        logging.info("Data appended successfully to logs.xlsx.")
        return True, None
    except Exception as e:
        logging.error("Error appending log: %s", e)
        return False, str(e)

@app.route("/", methods=["GET"])
def index():
    projects = load_projects()
    current_date = datetime.now().strftime("%Y-%m-%d")
    return render_template("form.html", projects=projects, current_date=current_date)

@app.route("/submit", methods=["POST"])
def submit():
    try:
        data = {
            'employee_id': request.form.get("employee_id", "").strip(),
            'employee_name': request.form.get("employee_name", "").strip(),
            'date': request.form.get("date"),
            'project': request.form.get("project", "").strip(),
            'hours': request.form.get("hours"),
            'description': request.form.get("description", "").strip()
        }

        anomalies = check_data_integrity(data)
        anomaly_message = "; ".join(anomalies) if anomalies else ""
        
        success, err_msg = append_log(data, anomaly_message)
        projects = load_projects()
        
        if success:
            # Always show a generic success message, even if anomalies exist
            return render_template("form.html", projects=projects, current_date=data['date'],
                                   message="Log submitted successfully!")
        else:
            return render_template("form.html", projects=projects, current_date=data['date'],
                                   error=f"Error saving log: {err_msg}")
    except Exception as e:
        logging.error("Error in submit route: %s", e)
        projects = load_projects()
        return render_template("form.html", projects=projects,
                               current_date=datetime.now().strftime("%Y-%m-%d"),
                               error="An unexpected error occurred.")

if __name__ == "__main__":
    app.run(debug=True)
