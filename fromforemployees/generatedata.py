import random
import os
from openpyxl import Workbook
from datetime import datetime, timedelta

# Same enumerations as in app.py for logs
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

# Possible TMs (Team Members) for logs/leaves
TEAM_MEMBERS = ["TM1", "TM2", "TM3", "TM4", "TM5"]

# Sample leave types
LEAVE_TYPES = ["Sick Leave", "Vacation", "Personal", "Training", "Other"]

def random_date_in_feb_2025():
    """Generate a random date in February 2025."""
    start = datetime(2025, 2, 1)
    end = datetime(2025, 2, 28)
    delta = end - start
    random_days = random.randrange(delta.days + 1)
    return (start + timedelta(days=random_days)).date()

def generate_logs(filename="logs.xlsx", num_records=10):
    """
    Generate a logs.xlsx with the 17 columns + 'Anomaly Reason',
    mirroring the fields from the new 17-field form.
    """
    wb = Workbook()
    sheet = wb.active
    sheet.title = "Logs"
    headers = [
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
    ]
    sheet.append(headers)

    for _ in range(num_records):
        # Randomly pick a date in Feb 2025, then adjust to Friday
        rand_date = random_date_in_feb_2025()
        while rand_date.weekday() != 4:  # 4 = Friday
            rand_date += timedelta(days=1)
            if rand_date.month != 2:  # if we exit February, break
                rand_date = datetime(2025, 2, 28).date()
                break
        status_date_str = rand_date.isoformat()

        main_project = random.choice(["MainProjA", "MainProjB", "MainProjC"])
        project_name = f"Project_{random.randint(100,999)}"
        milestones = "Milestone 1, Milestone 2"
        tm_value = random.choice(TEAM_MEMBERS)  # "TM1", "TM2", etc.

        # Start & Completion dates
        start_d = datetime(2025, 2, random.randint(1, 10)).date()
        completion_d = datetime(2025, 2, random.randint(11, 28)).date()
        if start_d > completion_d:
            start_d, completion_d = completion_d, start_d  # swap

        percent_completion = round(random.uniform(0, 100), 2)
        status = random.choice(["On Track", "At Risk", "Delayed", "Completed"])
        weekly_time_spent = round(random.uniform(0, 20), 2)
        projected_hours = round(random.uniform(5, 20), 2)

        f_area = random.choice(VALID_FUNCTIONAL_AREAS)
        p_cat = random.choice(VALID_PROJECT_CATEGORIES)
        comp = random.choice(VALID_COMPLEXITY)
        nov = random.choice(VALID_NOVELTY)
        out_type = random.choice(VALID_OUTPUT_TYPES)
        imp_type = random.choice(VALID_IMPACT_TYPES)

        row = [
            status_date_str,          # Status Date
            main_project,             # Main Project
            project_name,             # Project Name
            milestones,               # Key Milestones
            tm_value,                 # TM
            start_d.isoformat(),      # Start Date
            completion_d.isoformat(), # Completion Date
            percent_completion,       # % of Completion
            status,                   # Status
            weekly_time_spent,        # Weekly Time Spent
            projected_hours,          # Projected Hours
            f_area,                   # Functional Area
            p_cat,                    # Project Category
            comp,                     # Complexity
            nov,                      # Novelty
            out_type,                 # Output Type
            imp_type,                 # Impact Type
            ""                        # Anomaly Reason
        ]
        sheet.append(row)

    wb.save(filename)
    print(f"Generated {num_records} records in {filename}.")

def generate_leaves(filename="leaves.xlsx", num_leaves=10):
    """
    Generate a leaves.xlsx with columns:
    [TM, Date, Leave Type]
    We'll assume TM is the same field used in logs.xlsx.
    """
    wb = Workbook()
    sheet = wb.active
    sheet.title = "Leaves"
    headers = ["TM", "Date", "Leave Type"]
    sheet.append(headers)

    for _ in range(num_leaves):
        tm_value = random.choice(TEAM_MEMBERS)
        date_value = random_date_in_feb_2025().isoformat()
        leave_type = random.choice(LEAVE_TYPES)
        row = [tm_value, date_value, leave_type]
        sheet.append(row)

    wb.save(filename)
    print(f"Generated {num_leaves} records in {filename}.")

if __name__ == "__main__":
    # Generate both logs and leaves
    generate_logs(num_records=15)
    generate_leaves(num_leaves=8)