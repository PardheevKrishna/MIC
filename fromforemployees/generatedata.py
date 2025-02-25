import random
import os
from openpyxl import Workbook
from datetime import datetime, timedelta

# Enumerated lists for random data generation
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
VALID_NOVELTY = ["BAU repetitive", "One time repetitive", "New one time"]
VALID_OUTPUT_TYPES = [
    "Core production work",
    "Ad-hoc long-term projects",
    "Ad-hoc short-term projects",
    "Business Management",
    "Administration",
    "Trainings/L&D activities",
    "Others"
]
VALID_IMPACT_TYPES = ["Customer Experience", "Financial impact", "Insights", "Risk reduction", "Others"]

TEAM_MEMBERS = ["TM1", "TM2", "TM3", "TM4", "TM5"]
LEAVE_TYPES = ["Sick Leave", "Vacation", "Personal", "Training", "Other"]

def random_date_in_feb_2025():
    """Generate a random date in February 2025."""
    start = datetime(2025, 2, 1)
    end = datetime(2025, 2, 28)
    delta = end - start
    random_days = random.randrange(delta.days + 1)
    return (start + timedelta(days=random_days)).date()

def generate_logs(filename="logs.xlsx", num_records=15):
    """Generate logs.xlsx with 17 fields plus an 'Anomaly Reason' column."""
    wb = Workbook()
    sheet = wb.active
    sheet.title = "Logs"
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
    for _ in range(num_records):
        # Randomly choose a date in Feb 2025 and adjust to Friday
        rand_date = random_date_in_feb_2025()
        while rand_date.weekday() != 4:
            rand_date += timedelta(days=1)
            if rand_date.month != 2:
                rand_date = datetime(2025, 2, 28).date()
                break
        status_date_str = rand_date.isoformat()
        main_project = random.choice(["MainProjA", "MainProjB", "MainProjC"])
        project_name = f"Project_{random.randint(100,999)}"
        milestones = "Milestone 1, Milestone 2"
        tm = random.choice(TEAM_MEMBERS)
        start_d = datetime(2025, 2, random.randint(1, 10)).date()
        completion_d = datetime(2025, 2, random.randint(11, 28)).date()
        if start_d > completion_d:
            start_d, completion_d = completion_d, start_d
        pct = round(random.uniform(0, 100), 2)
        status = random.choice(["On Track", "At Risk", "Delayed", "Completed"])
        weekly_time_spent = round(random.uniform(0, 45), 2)
        projected_hours = round(random.uniform(5, 45), 2)
        f_area = random.choice(VALID_FUNCTIONAL_AREAS)
        p_cat = random.choice(VALID_PROJECT_CATEGORIES)
        comp = random.choice(VALID_COMPLEXITY)
        nov = random.choice(VALID_NOVELTY)
        out_type = random.choice(VALID_OUTPUT_TYPES)
        imp_type = random.choice(VALID_IMPACT_TYPES)
        row = [
            status_date_str,
            main_project,
            project_name,
            milestones,
            tm,
            start_d.isoformat(),
            completion_d.isoformat(),
            pct,
            status,
            weekly_time_spent,
            projected_hours,
            f_area,
            p_cat,
            comp,
            nov,
            out_type,
            imp_type,
            ""  # Anomaly Reason left empty
        ]
        sheet.append(row)
    wb.save(filename)
    print(f"Generated {num_records} records in {filename}.")

def generate_leaves(filename="leaves.xlsx", num_records=8):
    """Generate leaves.xlsx with columns: Team Member (TM), Date, Leave Type."""
    wb = Workbook()
    sheet = wb.active
    sheet.title = "Leaves"
    headers = ["Team Member (TM)", "Date", "Leave Type"]
    sheet.append(headers)
    for _ in range(num_records):
        tm = random.choice(TEAM_MEMBERS)
        date_str = random_date_in_feb_2025().isoformat()
        leave_type = random.choice(LEAVE_TYPES)
        sheet.append([tm, date_str, leave_type])
    wb.save(filename)
    print(f"Generated {num_records} records in {filename}.")

if __name__ == "__main__":
    generate_logs(num_records=15)
    generate_leaves(num_records=8)