import pandas as pd
import datetime
import re
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# ---------------------------
# Step 1: Read the Excel File
# ---------------------------
input_file = "input.xlsx"
df_data = pd.read_excel(input_file, sheet_name="Data")
df_control = pd.read_excel(input_file, sheet_name="Control")  # Read Control sheet

# Convert filemonth_dt to datetime (assuming mm/dd/yyyy format)
df_data['filemonth_dt'] = pd.to_datetime(df_data['filemonth_dt'], format='%m/%d/%Y')

# ---------------------------
# Step 2: Define the Two Dates
# ---------------------------
# Date1 is January 1, 2025, and Date2 is one month before (December 1, 2024)
date1 = datetime.datetime(2025, 1, 1)
date2 = datetime.datetime(2024, 12, 1)

# ---------------------------
# Step 3: Get Unique Sorted Field Names
# ---------------------------
fields = sorted(df_data['field_name'].unique())

# ---------------------------------------------------
# Step 4: Define Phrases (with escaped parentheses)
# ---------------------------------------------------
phrases = [
    "1\\)   F6CF Loan - Both Pop, Diff Values",
    "2\\)   CF Loan - Prior Null, Current Pop",
    "3\\)   CF Loan - Prior Pop, Current Null"
]

# Define a function that uses regex search for the phrases.
def contains_phrase(text, phrases):
    for phrase in phrases:
        if re.search(phrase, text):
            return True
    return False

# ---------------------------------------------------
# Step 5: Compute the Summary Aggregations per Field
# ---------------------------------------------------
summary_data = []
for field in fields:
    # For Missing Values: analysis_type is 'value_dist' and value_label contains "Missing"
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
    
    # For Month-to-Month Value Differences: analysis_type is 'pop_comp' and value_label contains any of the three phrases.
    mask_m2m_date1 = (
        (df_data['analysis_type'] == 'pop_comp') &
        (df_data['field_name'] == field) &
        (df_data['filemonth_dt'] == date1) &
        (df_data['value_label'].apply(lambda x: contains_phrase(x, phrases)))
    )
    m2m_sum_date1 = df_data.loc[mask_m2m_date1, 'value_records'].sum()
    
    mask_m2m_date2 = (
        (df_data['analysis_type'] == 'pop_comp') &
        (df_data['field_name'] == field) &
        (df_data['filemonth_dt'] == date2) &
        (df_data['value_label'].apply(lambda x: contains_phrase(x, phrases)))
    )
    m2m_sum_date2 = df_data.loc[mask_m2m_date2, 'value_records'].sum()
    
    summary_data.append([field, missing_sum_date1, missing_sum_date2, m2m_sum_date1, m2m_sum_date2])

# ---------------------------------------------------
# Step 6: Load Existing Workbook and Add Summary Sheet
# ---------------------------------------------------
# This will keep the original Data and Control sheets.
wb = load_workbook(input_file)

# Remove an existing Summary sheet if present.
if "Summary" in wb.sheetnames:
    ws_old = wb["Summary"]
    wb.remove(ws_old)

ws_summary = wb.create_sheet("Summary")

# ---------------------------------------------------
# Step 7: Define Styles for Headers and Cells
# ---------------------------------------------------
# Header style: blue fill with white bold text, centered
header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
header_font = Font(color="FFFFFF", bold=True)
center_alignment = Alignment(horizontal="center", vertical="center")

# Alternate row fills: white and light gray
white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
gray_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")

# Define a thick border style for cells with values.
thick_side = Side(border_style="thick", color="000000")
thick_border = Border(left=thick_side, right=thick_side, top=thick_side, bottom=thick_side)

# ---------------------------------------------------
# Step 8: Write Headers in the Summary Sheet
# ---------------------------------------------------
# Row 1: Title spanning A1:I1
ws_summary.merge_cells("A1:I1")
ws_summary["A1"] = "BDCOMM FRY14M Field Analysis Summary"
ws_summary["A1"].fill = header_fill
ws_summary["A1"].font = header_font
ws_summary["A1"].alignment = center_alignment

# Rows 2 and 3 for column headings:
# Merge A2:A3 for "Field Name"
ws_summary.merge_cells("A2:A3")
ws_summary["A2"] = "Field Name"
ws_summary["A2"].fill = header_fill
ws_summary["A2"].font = header_font
ws_summary["A2"].alignment = center_alignment

# Merge C2:D2 for "Missing Values"
ws_summary.merge_cells("C2:D2")
ws_summary["C2"] = "Missing Values"
ws_summary["C2"].fill = header_fill
ws_summary["C2"].font = header_font
ws_summary["C2"].alignment = center_alignment

# Row 3: Dates for Missing Values columns
ws_summary["C3"] = date1
ws_summary["C3"].number_format = "mm/dd/yyyy"
ws_summary["C3"].alignment = center_alignment

ws_summary["D3"] = date2
ws_summary["D3"].number_format = "mm/dd/yyyy"
ws_summary["D3"].alignment = center_alignment

# Merge F2:G2 for "Month to Month Value Differences"
ws_summary.merge_cells("F2:G2")
ws_summary["F2"] = "Month to Month Value Differences"
ws_summary["F2"].fill = header_fill
ws_summary["F2"].font = header_font
ws_summary["F2"].alignment = center_alignment

# Row 3: Dates for Month-to-Month columns
ws_summary["F3"] = date1
ws_summary["F3"].number_format = "mm/dd/yyyy"
ws_summary["F3"].alignment = center_alignment

ws_summary["G3"] = date2
ws_summary["G3"].number_format = "mm/dd/yyyy"
ws_summary["G3"].alignment = center_alignment

# Merge I2:I3 for "Approval Comments"
ws_summary.merge_cells("I2:I3")
ws_summary["I2"] = "Approval Comments"
ws_summary["I2"].fill = header_fill
ws_summary["I2"].font = header_font
ws_summary["I2"].alignment = center_alignment

# ---------------------------------------------------
# Step 9: Write the Computed Summary Data (starting row 4)
# ---------------------------------------------------
start_row = 4
for idx, row_data in enumerate(summary_data, start=start_row):
    # Alternate row fill: white for even, gray for odd (relative to summary data rows)
    fill_color = white_fill if (idx - start_row) % 2 == 0 else gray_fill

    # Column A: Field Name
    cell = ws_summary.cell(row=idx, column=1, value=row_data[0])
    cell.fill = fill_color
    cell.alignment = center_alignment
    if cell.value is not None:
        cell.border = thick_border

    # Column C and D: Missing Values for date1 and date2
    cell = ws_summary.cell(row=idx, column=3, value=row_data[1])
    cell.fill = fill_color
    cell.alignment = center_alignment
    if cell.value is not None:
        cell.border = thick_border

    cell = ws_summary.cell(row=idx, column=4, value=row_data[2])
    cell.fill = fill_color
    cell.alignment = center_alignment
    if cell.value is not None:
        cell.border = thick_border

    # Column F and G: Month-to-Month Value Differences for date1 and date2
    cell = ws_summary.cell(row=idx, column=6, value=row_data[3])
    cell.fill = fill_color
    cell.alignment = center_alignment
    if cell.value is not None:
        cell.border = thick_border

    cell = ws_summary.cell(row=idx, column=7, value=row_data[4])
    cell.fill = fill_color
    cell.alignment = center_alignment
    if cell.value is not None:
        cell.border = thick_border

# Optionally, adjust column widths for a neat report.
for col in ['A','B','C','D','E','F','G','H','I']:
    ws_summary.column_dimensions[col].width = 20

# ---------------------------------------------------
# Step 10: Save the Updated Workbook with all Sheets
# ---------------------------------------------------
output_file = "output.xlsx"
wb.save(output_file)