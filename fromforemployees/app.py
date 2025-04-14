import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import openpyxl
import os
import logging
from openpyxl.styles import PatternFill

# Set up logging
logging.basicConfig(level=logging.INFO)

LOGS_FILE = 'logs.xlsx'
LEAVES_FILE = 'leaves.xlsx'

# Orange fill for anomaly rows
ORANGE_FILL = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")

# Function to load suggestions from logs.xlsx
def load_suggestions():
    suggestions = {
        "main_project": set(),
        "project_name": set(),
        "project_key_milestones": set(),
        "tm": set(),
        "status": set()
    }
    if os.path.exists(LOGS_FILE):
        try:
            wb = openpyxl.load_workbook(LOGS_FILE)
            sheet = wb.active
            for row in sheet.iter_rows(min_row=2, values_only=True):
                if row[1]:
                    suggestions["main_project"].add(str(row[1]).strip())
                if row[2]:
                    suggestions["project_name"].add(str(row[2]).strip())
                if row[3]:
                    suggestions["project_key_milestones"].add(str(row[3]).strip())
                if row[4]:
                    suggestions["tm"].add(str(row[4]).strip())
                if row[8]:
                    suggestions["status"].add(str(row[8]).strip())
        except Exception as e:
            logging.error("Error loading suggestions: %s", e)
    for key in suggestions:
        suggestions[key] = sorted(list(suggestions[key]))
    return suggestions

# Function to handle custom input or selection with radio
def handle_radio_select(field_name, suggestions):
    # Create radio button to choose between custom input and dropdown
    option = st.radio(f"Choose {field_name}", ("Select from Dropdown", "Enter Custom Value"), key=field_name)
    
    # Show either text input or dropdown based on the radio choice
    if option == "Enter Custom Value":
        custom_value = st.text_input(f"Enter {field_name}", key=f"{field_name}_input")
        return custom_value
    else:
        selected_value = st.selectbox(f"Select {field_name}", options=suggestions, key=f"{field_name}_select")
        return selected_value

# Streamlit UI
st.set_page_config(page_title="Project Log", layout="wide")
st.title("Project Log")

# Load suggestions from the logs.xlsx
suggestions = load_suggestions()

# Form for input
with st.form(key="project_log_form"):
    status_date = st.date_input("Status Date (Every Friday)")

    # Handle the Main Project
    main_project = handle_radio_select("Main Project", suggestions["main_project"])
    
    # Handle Project Name
    project_name = handle_radio_select("Project Name", suggestions["project_name"])
    
    # Handle Project Key Milestones
    project_key_milestones = handle_radio_select("Project Key Milestones", suggestions["project_key_milestones"])
    
    # Handle Team Member (TM)
    tm = handle_radio_select("Team Member (TM)", suggestions["tm"])
    
    # Other fields
    start_date = st.date_input("Start Date")
    completion_date = st.date_input("Completion Date")
    percent_completion = st.number_input("% of Completion", min_value=0, max_value=100)
    status = st.selectbox("Status", options=suggestions["status"])
    weekly_time_spent = st.number_input("Weekly Time Spent (Hrs)", min_value=0.0, step=0.5)
    projected_hours = st.number_input("Projected Hours (E2E Implementation)", min_value=0.0, step=0.5)
    functional_area = st.selectbox("Functional Area", ["CRIT", "CRIT - Data Management", "CRIT - Data Governance", "CRIT - Regulatory Reporting", "CRIT - Portfolio Reporting", "CRIT - Transformation"])
    project_category = st.selectbox("Project Category", ["Data Infrastructure", "Monitoring & Insights", "Analytics / Strategy Development", "GDA Related", "Trainings and Team Meeting"])
    complexity = st.selectbox("Complexity (H, M, L)", ["H", "M", "L"])
    novelty = st.selectbox("Novelty", ["BAU repetitive", "One time repetitive", "New one time"])
    output_type = st.selectbox("Output Type", ["Core production work", "Ad-hoc long-term projects", "Ad-hoc short-term projects", "Business Management", "Administration", "Trainings/L&D activities", "Others"])
    impact_type = st.selectbox("Impact Type", ["Customer Experience", "Financial impact", "Insights", "Risk reduction", "Others"])

    submit_button = st.form_submit_button(label="Submit")

# Handle form submission
if submit_button:
    data = {
        "status_date": str(status_date),
        "main_project": main_project,
        "project_name": project_name,
        "project_key_milestones": project_key_milestones,
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

    # Validate and check for anomalies
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