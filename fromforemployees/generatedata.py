import random
from openpyxl import Workbook
from datetime import datetime, timedelta

def random_date(start, end):
    """Return a random datetime between start and end."""
    delta = end - start
    random_days = random.randrange(delta.days + 1)
    return start + timedelta(days=random_days)

def generate_logs(filename="logs.xlsx", num_logs=30):
    """
    Generate a random logs.xlsx file with the columns:
    Employee ID, Employee Name, Date, Project, Hours Worked, Description, Anomaly Reason
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Logs"
    headers = ["Employee ID", "Employee Name", "Date", "Project", "Hours Worked", "Description", "Anomaly Reason"]
    ws.append(headers)
    
    employees = [
        {"id": "E001", "name": "Alice"},
        {"id": "E002", "name": "Bob"},
        {"id": "E003", "name": "Charlie"},
        {"id": "E004", "name": "David"},
        {"id": "E005", "name": "Eva"}
    ]
    projects = ["Project Alpha", "Project Beta", "Project Gamma", "Form Maker", "Project Delta", "Project 1"]
    descriptions = [
        "Worked on feature X",
        "Completed module Y",
        "Reviewed code",
        "Fixed bugs",
        "Implemented design",
        "Tested application"
    ]
    
    # Date range for February 2025
    start_date = datetime(2025, 2, 1)
    end_date = datetime(2025, 2, 28)
    
    for _ in range(num_logs):
        emp = random.choice(employees)
        date = random_date(start_date, end_date).date().isoformat()
        project = random.choice(projects)
        hours = round(random.uniform(1, 10), 1)
        description = random.choice(descriptions)
        ws.append([emp["id"], emp["name"], date, project, hours, description, ""])
    
    wb.save(filename)
    print(f"Generated {filename} with {num_logs} log entries.")

def generate_leaves(filename="leaves.xlsx", num_leaves=10):
    """
    Generate a random leaves.xlsx file with the columns:
    Employee ID, Date, Leave Type
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Leaves"
    headers = ["Employee ID", "Date", "Leave Type"]
    ws.append(headers)
    
    employees = [
        {"id": "E001", "name": "Alice"},
        {"id": "E002", "name": "Bob"},
        {"id": "E003", "name": "Charlie"},
        {"id": "E004", "name": "David"},
        {"id": "E005", "name": "Eva"}
    ]
    leave_types = ["Sick Leave", "Vacation", "Personal"]
    
    # Date range for February 2025
    start_date = datetime(2025, 2, 1)
    end_date = datetime(2025, 2, 28)
    
    for _ in range(num_leaves):
        emp = random.choice(employees)
        date = random_date(start_date, end_date).date().isoformat()
        leave_type = random.choice(leave_types)
        ws.append([emp["id"], date, leave_type])
    
    wb.save(filename)
    print(f"Generated {filename} with {num_leaves} leave entries.")

if __name__ == "__main__":
    generate_logs()
    generate_leaves()
    