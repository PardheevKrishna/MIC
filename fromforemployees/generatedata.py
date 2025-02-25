import random
import os
from openpyxl import Workbook
from datetime import datetime, timedelta

# Same enumerations as in app.py
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

def random_date_in_feb_2025():
    """Generate a random date in February 2025."""
    start = datetime(2025, 2, 1)
    end = datetime(2025, 2, 28)
    delta = end - start
    random_days = random.randrange(delta.days + 1)
    return (start + timedelta(days=random_days)).date()

def generate_data(filename="logs.xlsx", num_records=10):
    """Generate a random logs.xlsx with the 17 columns + 'Anomaly Reason'."""
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
        # Randomly pick a Friday in February 2025 for the status date
        # We'll just pick a random date and then move to the nearest Friday
        rand_date = random_date_in_feb_2025()
        # Force it to be a Friday by adding or subtracting days
        # (In real usage, you'd ensure it’s truly a Friday.)
        # Here, let's just pick a Friday near the random date:
        while rand_date.weekday() != 4:
            rand_date += timedelta(days=1)
            if rand_date.month != 2:  # if we exit February, break
                rand_date = datetime(2025, 2, 28).date()
                break

        status_date_str = rand_date.isoformat()

        main_project = random.choice(["MainProjA", "MainProjB", "MainProjC"])
        project_name = f"Project_{random.randint(100,999)}"
        milestones = "Milestone 1, Milestone 2"
        tm = random.choice(["TM1", "TM2", "TM3"])
        start_d = datetime(2025, 2, random.randint(1, 10)).date()
        completion_d = datetime(2025, 2, random.randint(11, 28)).date()
        if start_d > completion_d:
            start_d, completion_d = completion_d, start_d  # swap to keep them in order

        pc = round(random.uniform(0, 100), 2)  # % of completion
        status = random.choice(["On Track", "At Risk", "Delayed", "Completed"])
        wts = round(random.uniform(0, 20), 2)  # Weekly Time Spent
        ph = round(random.uniform(5, 20), 2)   # Projected Hours
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
            pc,
            status,
            wts,
            ph,
            f_area,
            p_cat,
            comp,
            nov,
            out_type,
            imp_type,
            ""  # Anomaly Reason left empty for now
        ]
        sheet.append(row)

    wb.save(filename)
    print(f"Generated {num_records} records in {filename}.")

if __name__ == "__main__":
    generate_data(num_records=15)