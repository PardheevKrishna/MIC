import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from dateutil.relativedelta import relativedelta
from datetime import datetime

# --- Parameters ---
input_file = 'input.xlsx'           # your input Excel file
output_file = 'output_with_summary.xlsx'  # output file with Summary sheet

# --- Read data from Excel ---
data_df = pd.read_excel(input_file, sheet_name='Data')
# We also have a Control sheet but it is not used in the summary per instructions:
# control_df = pd.read_excel(input_file, sheet_name='Control')

# Convert filemonth_dt to datetime
data_df['filemonth_dt'] = pd.to_datetime(data_df['filemonth_dt'], format='%m/%d/%Y', errors='coerce')

# --- Define the two dates ---
# We set the primary date as Jan 1, 2024 and the second as one month before.
date1 = pd.Timestamp('2024-01-01')
# Subtract one month – this yields December 1, 2023
date2 = date1 - pd.DateOffset(months=1)

# --- Compute Missing Values counts for analysis_type 'value_dist' ---
missing_date1 = (
    data_df[(data_df['analysis_type'] == 'value_dist') &
            (data_df['filemonth_dt'] == date1) &
            (data_df['value_label'].str.contains('Missing', na=False))]
    .groupby('field_name')
    .size()
)
missing_date2 = (
    data_df[(data_df['analysis_type'] == 'value_dist') &
            (data_df['filemonth_dt'] == date2) &
            (data_df['value_label'].str.contains('Missing', na=False))]
    .groupby('field_name')
    .size()
)

# --- Compute Month-to-Month Value Differences sums for analysis_type 'pop_comp' ---
# The rows to be summed should have value_label containing any one of three phrases.
phrases = [
    "1)   F6CF Loan - Both Pop, Diff Values",
    "2)   CF Loan - Prior Null, Current Pop",
    "3)   CF Loan - Prior Pop, Current Null"
]
pattern = '|'.join(phrases)

popcomp_date1 = (
    data_df[(data_df['analysis_type'] == 'pop_comp') &
            (data_df['filemonth_dt'] == date1) &
            (data_df['value_label'].str.contains(pattern, na=False))]
    .groupby('field_name')['value_records']
    .sum()
)
popcomp_date2 = (
    data_df[(data_df['analysis_type'] == 'pop_comp') &
            (data_df['filemonth_dt'] == date2) &
            (data_df['value_label'].str.contains(pattern, na=False))]
    .groupby('field_name')['value_records']
    .sum()
)

# --- Get sorted unique field names ---
fields = sorted(data_df['field_name'].dropna().unique())

# --- Prepare summary data ---
# Each row will include:
#   Field Name, Missing count for date1, Missing count for date2,
#   Pop_Comp sum for date1, Pop_Comp sum for date2, and blank Approval Comments.
summary_rows = []
for field in fields:
    summary_rows.append({
        "Field Name": field,
        "Missing_2024_01_01": int(missing_date1.get(field, 0)),
        "Missing_2023_12_01": int(missing_date2.get(field, 0)),
        "PopComp_2024_01_01": float(popcomp_date1.get(field, 0)),
        "PopComp_2023_12_01": float(popcomp_date2.get(field, 0)),
        "Approval Comments": ""
    })

# --- Write Summary sheet using openpyxl ---
# Load the workbook
wb = load_workbook(input_file)
# Remove existing 'Summary' sheet if it exists
if 'Summary' in wb.sheetnames:
    del wb['Summary']
ws = wb.create_sheet('Summary')

# Define header fill and font (using a blue color and white text)
header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
header_font = Font(color="FFFFFF", bold=True)
center_align = Alignment(horizontal="center", vertical="center")

# --- Create the Title Row ---
ws.merge_cells('A1:I1')
ws['A1'] = "BDCOMM FRY14M Field Analysis Summary"
ws['A1'].fill = header_fill
ws['A1'].font = header_font
ws['A1'].alignment = center_align

# --- Setup the Multi-level Headers ---
# Column A: "Filed Name" spanning A2:A3
ws.merge_cells('A2:A3')
ws['A2'] = "Filed Name"
ws['A2'].fill = header_fill
ws['A2'].font = header_font
ws['A2'].alignment = center_align

# Columns C-D: "Missing Values" header spanning C2:D2
ws.merge_cells('C2:D2')
ws['C2'] = "Missing Values"
ws['C2'].fill = header_fill
ws['C2'].font = header_font
ws['C2'].alignment = center_align
# Set the sub-headers with dates (formatted as mm/dd/yyyy)
ws['C3'] = date1.strftime("%m/%d/%Y")
ws['D3'] = date2.strftime("%m/%d/%Y")
for cell in ['C3', 'D3']:
    ws[cell].fill = header_fill
    ws[cell].font = header_font
    ws[cell].alignment = center_align

# Columns F-G: "Month to Month Value Differences" header spanning F2:G2
ws.merge_cells('F2:G2')
ws['F2'] = "Month to Month Value Differences"
ws['F2'].fill = header_fill
ws['F2'].font = header_font
ws['F2'].alignment = center_align
ws['F3'] = date1.strftime("%m/%d/%Y")
ws['G3'] = date2.strftime("%m/%d/%Y")
for cell in ['F3', 'G3']:
    ws[cell].fill = header_fill
    ws[cell].font = header_font
    ws[cell].alignment = center_align

# Column I: "Approval Comments" header spanning I2:I3
ws.merge_cells('I2:I3')
ws['I2'] = "Approval Comments"
ws['I2'].fill = header_fill
ws['I2'].font = header_font
ws['I2'].alignment = center_align

# --- Write the data rows ---
# We'll start writing the field names and computed values from row 4 downward.
current_row = 4
for row in summary_rows:
    # Column A: Field Name
    ws.cell(row=current_row, column=1, value=row["Field Name"])
    # Column C: Missing count for date1
    ws.cell(row=current_row, column=3, value=row["Missing_2024_01_01"])
    # Column D: Missing count for date2
    ws.cell(row=current_row, column=4, value=row["Missing_2023_12_01"])
    # Column F: PopComp sum for date1
    ws.cell(row=current_row, column=6, value=row["PopComp_2024_01_01"])
    # Column G: PopComp sum for date2
    ws.cell(row=current_row, column=7, value=row["PopComp_2023_12_01"])
    # Column I: Approval Comments (left blank)
    ws.cell(row=current_row, column=9, value=row["Approval Comments"])
    current_row += 1

# --- Save the workbook with the new Summary sheet ---
wb.save(output_file)
print(f"Summary sheet created and saved in {output_file}")