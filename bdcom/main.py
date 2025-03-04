import pandas as pd
import datetime
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment

# --- Step 1. Read the Data ---
# Read the 'Data' sheet from the Excel file.
df_data = pd.read_excel("input.xlsx", sheet_name="Data")

# Convert filemonth_dt to datetime (assuming format mm/dd/yyyy)
df_data['filemonth_dt'] = pd.to_datetime(df_data['filemonth_dt'], format='%m/%d/%Y')

# (Control sheet is not used in our summary calculations per instructions.)
# df_control = pd.read_excel("input.xlsx", sheet_name="Control")

# --- Step 2. Define Dates and Get Unique Field Names ---
# Define the two dates: Jan 1, 2024 and one month before (Dec 1, 2023)
date1 = datetime.datetime(2024, 1, 1)
date2 = datetime.datetime(2023, 12, 1)

# Get the sorted list of unique field names from the Data sheet
fields = sorted(df_data['field_name'].unique())

# --- Step 3. Compute Aggregated Sums for Each Field ---
summary_data = []
# For month-to-month differences, define the three phrases to check
phrases = [
    "1)   F6CF Loan - Both Pop, Diff Values", 
    "2)   CF Loan - Prior Null, Current Pop", 
    "3)   CF Loan - Prior Pop, Current Null"
]

for field in fields:
    # Missing Values: analysis_type 'value_dist' and value_label contains "Missing"
    mask_missing_date1 = (
        (df_data['analysis_type'] == 'value_dist') &
        (df_data['field_name'] == field) &
        (df_data['filemonth_dt'] == date1) &
        (df_data['value_label'].str.contains("Missing", case=False, na=False))
    )
    missing_sum_date1 = df_data.loc[mask_missing_date1, 'value_records'].sum()
    
    mask_missing_date2 = (
        (df_data['analysis_type'] == 'value_dist') &
        (df_data['field_name'] == field) &
        (df_data['filemonth_dt'] == date2) &
        (df_data['value_label'].str.contains("Missing", case=False, na=False))
    )
    missing_sum_date2 = df_data.loc[mask_missing_date2, 'value_records'].sum()
    
    # Month-to-Month Value Differences: analysis_type 'pop_comp' and value_label contains any of the three phrases.
    mask_m2m_date1 = (
        (df_data['analysis_type'] == 'pop_comp') &
        (df_data['field_name'] == field) &
        (df_data['filemonth_dt'] == date1) &
        (df_data['value_label'].apply(lambda x: any(phrase in x for phrase in phrases)))
    )
    m2m_sum_date1 = df_data.loc[mask_m2m_date1, 'value_records'].sum()
    
    mask_m2m_date2 = (
        (df_data['analysis_type'] == 'pop_comp') &
        (df_data['field_name'] == field) &
        (df_data['filemonth_dt'] == date2) &
        (df_data['value_label'].apply(lambda x: any(phrase in x for phrase in phrases)))
    )
    m2m_sum_date2 = df_data.loc[mask_m2m_date2, 'value_records'].sum()
    
    # Append the computed values for the current field.
    # The list order is: field, missing_sum_date1, missing_sum_date2, m2m_sum_date1, m2m_sum_date2
    summary_data.append([field, missing_sum_date1, missing_sum_date2, m2m_sum_date1, m2m_sum_date2])

# --- Step 4. Create the Summary Sheet with openpyxl ---
# Create a new workbook and select the active sheet, renaming it to "Summary".
wb = Workbook()
ws = wb.active
ws.title = "Summary"

# Define header formatting: a blue fill with white bold text.
header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
header_font = Font(color="FFFFFF", bold=True)
center_alignment = Alignment(horizontal="center", vertical="center")

# Row 1: Merge A1:I1 and set title.
ws.merge_cells("A1:I1")
ws["A1"] = "BDCOMM FRY14M Field Analysis Summary"
ws["A1"].fill = header_fill
ws["A1"].font = header_font
ws["A1"].alignment = center_alignment

# Row 2 and 3 Headers:
# Merge A2:A3 for "Filed Name"
ws.merge_cells("A2:A3")
ws["A2"] = "Filed Name"
ws["A2"].fill = header_fill
ws["A2"].font = header_font
ws["A2"].alignment = center_alignment

# Merge C2:D2 for "Missing Values"
ws.merge_cells("C2:D2")
ws["C2"] = "Missing Values"
ws["C2"].fill = header_fill
ws["C2"].font = header_font
ws["C2"].alignment = center_alignment

# In row 3, place the dates under Missing Values.
ws["C3"] = date1
ws["C3"].number_format = "mm/dd/yyyy"
ws["C3"].alignment = center_alignment

ws["D3"] = date2
ws["D3"].number_format = "mm/dd/yyyy"
ws["D3"].alignment = center_alignment

# Merge F2:G2 for "Month to Month Value Differences"
ws.merge_cells("F2:G2")
ws["F2"] = "Month to Month Value Differences"
ws["F2"].fill = header_fill
ws["F2"].font = header_font
ws["F2"].alignment = center_alignment

# In row 3, place the dates under Month to Month Value Differences.
ws["F3"] = date1
ws["F3"].number_format = "mm/dd/yyyy"
ws["F3"].alignment = center_alignment

ws["G3"] = date2
ws["G3"].number_format = "mm/dd/yyyy"
ws["G3"].alignment = center_alignment

# Merge I2:I3 for "Approval Comments"
ws.merge_cells("I2:I3")
ws["I2"] = "Approval Comments"
ws["I2"].fill = header_fill
ws["I2"].font = header_font
ws["I2"].alignment = center_alignment

# --- Step 5. Write the Computed Data ---
# Start writing the data from row 4.
start_row = 4
for idx, row_data in enumerate(summary_data, start=start_row):
    # Write the field name in column A.
    ws.cell(row=idx, column=1, value=row_data[0])
    # Write Missing Values sums: date1 goes in column C, date2 in column D.
    ws.cell(row=idx, column=3, value=row_data[1])
    ws.cell(row=idx, column=4, value=row_data[2])
    # Write Month-to-Month Value Differences: date1 in column F, date2 in column G.
    ws.cell(row=idx, column=6, value=row_data[3])
    ws.cell(row=idx, column=7, value=row_data[4])
    # Columns B, E, H, and I (Approval Comments) are left blank.

# Optionally adjust column widths for a better view.
for col in ['A','B','C','D','E','F','G','H','I']:
    ws.column_dimensions[col].width = 20

# --- Step 6. Save the Output Workbook ---
wb.save("output.xlsx")