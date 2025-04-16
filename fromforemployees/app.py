import streamlit as st
import sqlite3
import logging
from datetime import datetime, timedelta

# Set up logging
logging.basicConfig(level=logging.INFO)

# Database file name
DB_FILE = "logs.db"

# Predefined team members; these will be used in the dropdown for TM.
EMPLOYEES = ["Alice", "Bob", "Carol", "Dave"]

# -------------------------------
# Initialize the SQLite database and create table if not exists.
def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            status_date TEXT,
            main_project TEXT,
            project_name TEXT,
            project_key_milestones TEXT,
            tm TEXT,
            start_date TEXT,
            completion_date TEXT,
            percent_completion REAL,
            status TEXT,
            weekly_time_spent REAL,
            projected_hours REAL,
            functional_area TEXT,
            project_category TEXT,
            complexity TEXT,
            novelty TEXT,
            output_type TEXT,
            impact_type TEXT,
            anomaly_reason TEXT
        )
    """)
    conn.commit()
    conn.close()

# -------------------------------
# Insert a log record into the database.
def insert_log(data, anomaly_reason):
    try:
        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        c.execute("""
          INSERT INTO logs (
              status_date, main_project, project_name, project_key_milestones, tm, 
              start_date, completion_date, percent_completion, status, weekly_time_spent, 
              projected_hours, functional_area, project_category, complexity, novelty, output_type, impact_type, anomaly_reason
          ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
              data["status_date"],
              data["main_project"],
              data["project_name"],
              data["project_key_milestones"],
              data["tm"],
              data["start_date"],
              data["completion_date"],
              float(data["percent_completion"]),
              data["status"],
              float(data["weekly_time_spent"]),
              float(data["projected_hours"]),
              data["functional_area"],
              data["project_category"],
              data["complexity"],
              data["novelty"],
              data["output_type"],
              data["impact_type"],
              anomaly_reason
        ))
        conn.commit()
        conn.close()
        logging.info("Log inserted successfully for %s.", data["tm"])
        return True, None
    except Exception as e:
        logging.error("Error inserting log: %s", e)
        return False, str(e)

# -------------------------------
# Load suggestions for dropdowns by querying distinct values in the logs table.
def load_suggestions():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT DISTINCT main_project FROM logs WHERE main_project IS NOT NULL AND main_project!=''")
    main_projects = sorted([row[0] for row in c.fetchall()])
    c.execute("SELECT DISTINCT project_name FROM logs WHERE project_name IS NOT NULL AND project_name!=''")
    project_names = sorted([row[0] for row in c.fetchall()])
    c.execute("SELECT DISTINCT project_key_milestones FROM logs WHERE project_key_milestones IS NOT NULL AND project_key_milestones!=''")
    project_key_milestones = sorted([row[0] for row in c.fetchall()])
    conn.close()
    return {
        "main_project": main_projects,
        "project_name": project_names,
        "project_key_milestones": project_key_milestones
    }

# -------------------------------
# Validate the provided data.
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
        completion_date = datetime.strptime(data['completion_date'], "%Y-%m-%d").date()
        start_date = datetime.strptime(data['start_date'], "%Y-%m-%d").date()
        if start_date > completion_date:
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

# -------------------------------
# Count leave days for a team member.
# (This demo function always returns 0; replace with actual logic if available.)
def count_leave_days(tm, status_date):
    return 0

# -------------------------------
# Check for anomalies based on Weekly Time Spent.
def check_anomalies(data):
    anomalies = []
    try:
        current_date = datetime.strptime(data['status_date'], "%Y-%m-%d").date()
        weekly_time_spent = float(data['weekly_time_spent'])
    except:
        return anomalies
    tm = data['tm'].strip()
    leave_days = count_leave_days(tm, data['status_date'])
    allowed_days = 5 - leave_days
    max_allowed_hours = allowed_days * 9
    if weekly_time_spent > max_allowed_hours:
        anomalies.append(f"Weekly Time Spent exceeds allowed limit. With {leave_days} leave day(s), maximum allowed is {max_allowed_hours} hrs.")
    return anomalies

# -------------------------------
# Initialize the database on startup.
init_db()

# -------------------------------
# STREAMLIT UI
st.set_page_config(page_title="Project Log", layout="wide")
st.title("Project Log")

# Load suggestions (these queries are fast in SQLite)
suggestions = load_suggestions()

# In the UI we allow the user to either type their value or choose one from a dropdown.
with st.form(key="project_log_form"):
    status_date = st.date_input("Status Date")
    
    # For Main project, Name of the Project, and Project Key Milestones,
    # we show a text input and a selectbox.
    main_project = st.text_input("Enter Main project (or choose below):", key="main_project_input")
    main_project_select = st.selectbox("Select Main project", options=[""] + suggestions["main_project"], key="main_project_select")
    main_project_val = main_project.strip() if main_project.strip() else main_project_select

    project_name = st.text_input("Enter Name of the Project (or choose below):", key="project_name_input")
    project_name_select = st.selectbox("Select Name of the Project", options=[""] + suggestions["project_name"], key="project_name_select")
    project_name_val = project_name.strip() if project_name.strip() else project_name_select

    project_key_milestones = st.text_input("Enter Project Key Milestones (or choose below):", key="project_key_milestones_input")
    project_key_milestones_select = st.selectbox("Select Project Key Milestones", options=[""] + suggestions["project_key_milestones"], key="project_key_milestones_select")
    project_key_milestones_val = project_key_milestones.strip() if project_key_milestones.strip() else project_key_milestones_select
    
    tm = st.selectbox("Team Member (TM)", options=EMPLOYEES)
    start_date = st.date_input("Start Date")
    completion_date = st.date_input("Completion Date")
    percent_completion = st.number_input("% of Completion", min_value=0, max_value=100)
    status = st.selectbox("Status", options=["Completed", "In Progress"])
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
        "project_key_milestones": project_key_milestones_val,
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
        success, err_msg = insert_log(data, anomaly_message)
        if success:
            st.success("Log submitted successfully!")
        else:
            st.error(f"Error saving log: {err_msg}")