import streamlit as st
from datetime import datetime, timedelta
import openpyxl
import os
import logging
from openpyxl.styles import PatternFill

# Set up logging
logging.basicConfig(level=logging.INFO)

# File name for the Excel workbook
LOGS_FILE = 'logs.xlsx'
LEAVES_FILE = 'leaves.xlsx'

# Predefined list of employees for the Team Member (TM) dropdown.
EMPLOYEES = ["Alice", "Bob", "Carol", "Dave"]

# Define the header row that must exist in each employee’s worksheet.
HEADERS = [
    "Status Date (Every Friday)", "Main project", "Name of the Project", "Project Key Milestones", "TM",
    "Start Date", "Completion Date % of Completion", "Status", "Weekly Time Spent(Hrs)",
    "Projected hours (Based on the Project: End to End implementation)",
    "Functional Area (CRIT, CRIT - Data Management, CRIT - Data Governance, CRIT - Regulatory Reporting, CRIT - Portfolio Reporting, CRIT - Transformation)",
    "Project Category (Data Infrastructure, Monitoring & Insights, Analytics / Strategy Development, GDA Related, Trainings and Team Meeting)",
    "Complexity (H,M,L)",
    "Novelty (BAU repetitive, One time repetitive, New one time)",
    "Output Type (Core production work, Ad-hoc long-term projects, Ad-hoc short-term projects, Business Management, Administration, Trainings/L&D activities, Others)",
    "Impact type (Customer Experience, Financial impact, Insights, Risk reduction, Others)",
    "Anomaly Reason"
]

# Orange fill for anomaly rows.
ORANGE_FILL = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")

# ==============================================================================
# Helper: Load (or create) the Excel workbook and cache it in session state.
def get_workbook():
    # Check if workbook is already loaded in session_state.
    if "workbook" not in st.session_state:
        if os.path.exists(LOGS_FILE):
            st.session_state.workbook = openpyxl.load_workbook(LOGS_FILE)
            logging.info("Workbook loaded from disk.")
        else:
            # Create new workbook and remove default sheet.
            wb = openpyxl.Workbook()
            default_sheet = wb.active
            wb.remove(default_sheet)
            st.session_state.workbook = wb
            logging.info("New workbook created.")
    return st.session_state.workbook

# ==============================================================================
# Helper: Load suggestions by scanning every employee’s sheet (if it exists)
# for values in the specified columns:
#  • Main project from column 2
#  • Name of the Project from column 3
#  • Project Key Milestones from column 4
def load_suggestions():
    suggestions = {
        "main_project": set(),
        "project_name": set(),
        "project_key_milestones": set()
    }
    wb = get_workbook()
    for sheet in wb.worksheets:
        # Process only sheets whose title is in the predefined EMPLOYEES.
        if sheet.title not in EMPLOYEES:
            continue
        # Iterate over rows in the sheet (skip header)
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if len(row) >= 2 and row[1]:
                suggestions["main_project"].add(str(row[1]).strip())
            if len(row) >= 3 and row[2]:
                suggestions["project_name"].add(str(row[2]).strip())
            if len(row) >= 4 and row[3]:
                suggestions["project_key_milestones"].add(str(row[3]).strip())
    # Sort the sets into lists.
    for key in suggestions:
        suggestions[key] = sorted(list(suggestions[key]))
    # For Status, we fix the dropdown to these two options.
    status_suggestions = ["Completed", "In Progress"]
    return suggestions, status_suggestions

# ==============================================================================
# Helper: Returns a value either from text input or dropdown (text input takes precedence).
def handle_input_or_select(field_name, suggestions):
    # Display both widgets.
    custom_val = st.text_input(f"Enter {field_name} (or choose from dropdown):", key=f"{field_name}_input")
    selected_val = st.selectbox(f"Select {field_name}", options=[""] + suggestions, key=f"{field_name}_select")
    return custom_val.strip() if custom_val.strip() else selected_val

# ==============================================================================
# Validation of the input data.
def validate_fields(data):
    errors = []
    try:
        datetime.strptime(data['status_date'], "%Y-%m-%d")
    except ValueError:
        errors.append("Invalid Status Date format.")
        
    if not data['main_project'].strip():
        errors.append("Main project is required.")
    if not data['project_name'].strip():
        errors.append("Name of the Project is required.")
    if not data['project_key_milestones'].strip():
        errors.append("Project Key Milestones are required.")
    if not data['tm'].strip():
        errors.append("Team Member (TM) is required.")
        
    try:
        datetime.strptime(data['start_date'], "%Y-%m-%d")
    except ValueError:
        errors.append("Invalid Start Date format.")
    try:
        comp_date = datetime.strptime(data['completion_date'], "%Y-%m-%d").date()
        start_date = datetime.strptime(data['start_date'], "%Y-%m-%d").date()
        if start_date > comp_date:
            errors.append("Start Date cannot be after Completion Date.")
    except ValueError:
        errors.append("Invalid Completion Date format.")
            
    try:
        pct = float(data['percent_completion'])
        if pct < 0 or pct > 100:
            errors.append("% of Completion must be between 0 and 100.")
    except ValueError:
        errors.append("Invalid % of Completion.")
        
    if not data['status'].strip():
        errors.append("Status is required.")
        
    try:
        wts = float(data['weekly_time_spent'])
        if wts < 0:
            errors.append("Weekly Time Spent cannot be negative.")
    except ValueError:
        errors.append("Invalid Weekly Time Spent value.")
        
    try:
        ph = float(data['projected_hours'])
        if ph < 0:
            errors.append("Projected Hours cannot be negative.")
    except ValueError:
        errors.append("Invalid Projected Hours value.")
        
    if not data['functional_area'].strip():
        errors.append("Functional Area is required.")
    if not data['project_category'].strip():
        errors.append("Project Category is required.")
    if data['complexity'] not in ["H", "M", "L"]:
        errors.append("Complexity must be H, M, or L.")
    if data['novelty'] not in ["BAU repetitive", "One time repetitive", "New one time"]:
        errors.append("Novelty must be one of BAU repetitive, One time repetitive, or New one time.")
    if not data['output_type'].strip():
        errors.append("Output Type is required.")
    if not data['impact_type'].strip():
        errors.append("Impact type is required.")
    return errors

# ==============================================================================
# (Optional) Count leave days for a team member over the week that includes status_date.
def count_leave_days(tm, status_date):
    # If you have actual logic to count leave days (by reading LEAVES_FILE etc.), add it here.
    # For now, we return 0.
    return 0

# ==============================================================================
# Check for anomalies based on weekly time spent.
def check_anomalies(data):
    anomalies = []
    try:
        current_date = datetime.strptime(data['status_date'], "%Y-%m-%d").date()
        weekly_hours = float(data['weekly_time_spent'])
    except:
        return anomalies
    tm = data['tm'].strip()
    leave_days = count_leave_days(tm, data['status_date'])
    allowed_days = 5 - leave_days
    max_hours = allowed_days * 9
    if weekly_hours > max_hours:
        anomalies.append(f"Weekly Time Spent exceeds allowed limit. With {leave_days} leave day(s), maximum allowed is {max_hours} hrs.")
    return anomalies

# ==============================================================================
# Append the new log as a row in the employee's sheet within the cached workbook.
def append_log(data, anomaly_message):
    wb = get_workbook()
    sheet_name = data['tm']
    # Create the sheet if it doesn't exist.
    if sheet_name not in wb.sheetnames:
        sheet = wb.create_sheet(title=sheet_name)
        sheet.append(HEADERS)
    else:
        sheet = wb[sheet_name]
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
    if anomaly_message:
        for col in range(1, 19):
            sheet.cell(row=new_row, column=col).fill = ORANGE_FILL
    try:
        wb.save(LOGS_FILE)
        logging.info("Log appended successfully for %s.", data['tm'])
        return True, None
    except Exception as e:
        logging.error("Error saving workbook: %s", e)
        return False, str(e)

# ==============================================================================
# STREAMLIT UI
st.set_page_config(page_title="Project Log", layout="wide")
st.title("Project Log")

# Load suggestions from the workbook
suggestions, status_suggestions = load_suggestions()

with st.form(key="project_log_form"):
    status_date = st.date_input("Status Date")
    
    # For these fields, the user can type or choose from suggestions.
    main_project_inp = st.text_input("Enter Main project (or choose below):", key="main_project_input")
    main_project_sel = st.selectbox("Select Main project", options=[""] + suggestions["main_project"], key="main_project_select")
    main_project_val = main_project_inp.strip() if main_project_inp.strip() else main_project_sel

    project_name_inp = st.text_input("Enter Name of the Project (or choose below):", key="project_name_input")
    project_name_sel = st.selectbox("Select Name of the Project", options=[""] + suggestions["project_name"], key="project_name_select")
    project_name_val = project_name_inp.strip() if project_name_inp.strip() else project_name_sel

    pkm_inp = st.text_input("Enter Project Key Milestones (or choose below):", key="pkm_input")
    pkm_sel = st.selectbox("Select Project Key Milestones", options=[""] + suggestions["project_key_milestones"], key="pkm_select")
    pkm_val = pkm_inp.strip() if pkm_inp.strip() else pkm_sel

    tm = st.selectbox("Team Member (TM)", options=EMPLOYEES)
    start_date = st.date_input("Start Date")
    completion_date = st.date_input("Completion Date")
    percent_completion = st.number_input("% of Completion", min_value=0, max_value=100)
    status = st.selectbox("Status", options=status_suggestions)
    weekly_time_spent = st.number_input("Weekly Time Spent (Hrs)", min_value=0.0, step=0.5)
    projected_hours = st.number_input("Projected hours (Based on the Project: End to End implementation)", min_value=0.0, step=0.5)
    functional_area = st.selectbox(
        "Functional Area (CRIT, CRIT - Data Management, CRIT - Data Governance, CRIT - Regulatory Reporting, CRIT - Portfolio Reporting, CRIT - Transformation)",
        options=["CRIT", "CRIT - Data Management", "CRIT - Data Governance", "CRIT - Regulatory Reporting", "CRIT - Portfolio Reporting", "CRIT - Transformation"]
    )
    project_category = st.selectbox(
        "Project Category (Data Infrastructure, Monitoring & Insights, Analytics / Strategy Development, GDA Related, Trainings and Team Meeting)",
        options=["Data Infrastructure", "Monitoring & Insights", "Analytics / Strategy Development", "GDA Related", "Trainings and Team Meeting"]
    )
    complexity = st.selectbox("Complexity (H,M,L)", options=["H", "M", "L"])
    novelty = st.selectbox("Novelty (BAU repetitive, One time repetitive, New one time)", options=["BAU repetitive", "One time repetitive", "New one time"])
    output_type = st.selectbox(
        "Output Type (Core production work, Ad-hoc long-term projects, Ad-hoc short-term projects, Business Management, Administration, Trainings/L&D activities, Others)",
        options=["Core production work", "Ad-hoc long-term projects", "Ad-hoc short-term projects", "Business Management", "Administration", "Trainings/L&D activities", "Others"]
    )
    impact_type = st.selectbox(
        "Impact type (Customer Experience, Financial impact, Insights, Risk reduction, Others)",
        options=["Customer Experience", "Financial impact", "Insights", "Risk reduction", "Others"]
    )

    submit_button = st.form_submit_button(label="Submit")

if submit_button:
    data = {
        "status_date": str(status_date),
        "main_project": main_project_val,
        "project_name": project_name_val,
        "project_key_milestones": pkm_val,
        "tm": tm,
        "start_date": str(start_date),
        "completion_date": str(completion_date),
        "percent_completion": str(percent_completion),
        "status": status,
        "weekly_time_spent": str(weekly_time_spent),
        "projected_hours": str(projected_hours),
        "functional_area": functional_area,
        "project_category": project_category,
        "complexity": complexity,
        "novelty": novelty,
        "output_type": output_type,
        "impact_type": impact_type
    }
    
    field_errors = validate_fields(data)
    if field_errors:
        st.error(" ".join(field_errors))
    else:
        anomalies = check_anomalies(data)
        anomaly_message = "; ".join(anomalies) if anomalies else ""
        success, err_msg = append_log(data, anomaly_message)
        if success:
            st.success("Log submitted successfully!")
        else:
            st.error(f"Error saving log: {err_msg}")