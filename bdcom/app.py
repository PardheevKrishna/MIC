import streamlit as st
import pandas as pd
import datetime
import os
import re
from io import BytesIO
from dateutil.relativedelta import relativedelta
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# ---------------------------------------------
# Sidebar: File selection and Date Input
# ---------------------------------------------
st.sidebar.header("File & Date Selection")

# List all Excel files in the current folder
folder_path = "."
files = [f for f in os.listdir(folder_path) if f.endswith(".xlsx")]
if not files:
    st.sidebar.error("No Excel (.xlsx) files found in the folder.")
else:
    selected_file = st.sidebar.selectbox("Select Excel File", files)

# Date selection for date1 (default set to Jan 1, 2025)
selected_date = st.sidebar.date_input("Select Date for Date1", datetime.date(2025, 1, 1))
# Convert selected_date (a date object) to a datetime object
date1 = datetime.datetime.combine(selected_date, datetime.datetime.min.time())
# date2 is one month before date1
date2 = date1 - relativedelta(months=1)

# ---------------------------------------------
# When the user clicks the button, process the file
# ---------------------------------------------
if st.sidebar.button("Generate Summary"):
    try:
        # Read the two sheets from the selected Excel file
        df_data = pd.read_excel(selected_file, sheet_name="Data")
        df_control = pd.read_excel(selected_file, sheet_name="Control")
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
    else:
        # Ensure the date column is a datetime
        df_data['filemonth_dt'] = pd.to_datetime(df_data['filemonth_dt'], format='%m/%d/%Y')

        # Get unique sorted field names
        fields = sorted(df_data['field_name'].unique())

        # Define the phrases (with escaped parentheses) for month-to-month differences
        phrases = [
            "1\\)   F6CF Loan - Both Pop, Diff Values",
            "2\\)   CF Loan - Prior Null, Current Pop",
            "3\\)   CF Loan - Prior Pop, Current Null"
        ]

        def contains_phrase(text, phrases):
            for phrase in phrases:
                if re.search(phrase, text):
                    return True
            return False

        # Compute the summary aggregations for each field
        summary_list = []
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

            # Month-to-Month Value Differences: analysis_type 'pop_comp' and value_label contains any of the phrases.
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

            summary_list.append([field, missing_sum_date1, missing_sum_date2, m2m_sum_date1, m2m_sum_date2])

        # Create a summary DataFrame with descriptive column names
        summary_df = pd.DataFrame(
            summary_list,
            columns=[
                "Field Name", 
                f"Missing Values ({date1.strftime('%m/%d/%Y')})", 
                f"Missing Values ({date2.strftime('%m/%d/%Y')})", 
                f"M2M Value Diff ({date1.strftime('%m/%d/%Y')})", 
                f"M2M Value Diff ({date2.strftime('%m/%d/%Y')})"
            ]
        )

        # ---------------------------------------------
        # Style the DataFrame with alternating row colors and thick cell borders
        # ---------------------------------------------
        def highlight_alternating_rows(row):
            # Alternate between white and light gray based on row index
            return ['background-color: white' if row.name % 2 == 0 else 'background-color: #D3D3D3'] * len(row)

        styled_df = summary_df.style.apply(highlight_alternating_rows, axis=1).set_table_styles(
            [{'selector': 'td', 'props': [('border', '2px solid black')]}]
        )

        # ---------------------------------------------
        # Display a title and the styled summary table
        # ---------------------------------------------
        st.markdown("<h1 style='text-align: center; color: #4F81BD;'>BDCOMM FRY14M Field Analysis Summary</h1>", unsafe_allow_html=True)
        st.markdown(styled_df.render(), unsafe_allow_html=True)

        # ---------------------------------------------
        # Create an updated Excel workbook including the existing sheets and a new Summary sheet
        # ---------------------------------------------
        wb = load_workbook(selected_file)
        # Remove any existing Summary sheet
        if "Summary" in wb.sheetnames:
            ws_old = wb["Summary"]
            wb.remove(ws_old)
        ws_summary = wb.create_sheet("Summary")

        # Define styles for headers and cells
        header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        center_alignment = Alignment(horizontal="center", vertical="center")
        white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        gray_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
        thick_side = Side(border_style="thick", color="000000")
        thick_border = Border(left=thick_side, right=thick_side, top=thick_side, bottom=thick_side)

        # Row 1: Title (merged A1:I1)
        ws_summary.merge_cells("A1:I1")
        ws_summary["A1"] = "BDCOMM FRY14M Field Analysis Summary"
        ws_summary["A1"].fill = header_fill
        ws_summary["A1"].font = header_font
        ws_summary["A1"].alignment = center_alignment

        # Row 2-3 headers:
        ws_summary.merge_cells("A2:A3")
        ws_summary["A2"] = "Field Name"
        ws_summary["A2"].fill = header_fill
        ws_summary["A2"].font = header_font
        ws_summary["A2"].alignment = center_alignment

        ws_summary.merge_cells("C2:D2")
        ws_summary["C2"] = "Missing Values"
        ws_summary["C2"].fill = header_fill
        ws_summary["C2"].font = header_font
        ws_summary["C2"].alignment = center_alignment

        ws_summary["C3"] = date1
        ws_summary["C3"].number_format = "mm/dd/yyyy"
        ws_summary["C3"].alignment = center_alignment

        ws_summary["D3"] = date2
        ws_summary["D3"].number_format = "mm/dd/yyyy"
        ws_summary["D3"].alignment = center_alignment

        ws_summary.merge_cells("F2:G2")
        ws_summary["F2"] = "Month to Month Value Differences"
        ws_summary["F2"].fill = header_fill
        ws_summary["F2"].font = header_font
        ws_summary["F2"].alignment = center_alignment

        ws_summary["F3"] = date1
        ws_summary["F3"].number_format = "mm/dd/yyyy"
        ws_summary["F3"].alignment = center_alignment

        ws_summary["G3"] = date2
        ws_summary["G3"].number_format = "mm/dd/yyyy"
        ws_summary["G3"].alignment = center_alignment

        ws_summary.merge_cells("I2:I3")
        ws_summary["I2"] = "Approval Comments"
        ws_summary["I2"].fill = header_fill
        ws_summary["I2"].font = header_font
        ws_summary["I2"].alignment = center_alignment

        # Write the computed summary data starting from row 4
        start_row_ws = 4
        for idx, row in enumerate(summary_list, start=start_row_ws):
            fill_color = white_fill if (idx - start_row_ws) % 2 == 0 else gray_fill

            cell = ws_summary.cell(row=idx, column=1, value=row[0])
            cell.fill = fill_color
            cell.alignment = center_alignment
            if cell.value is not None:
                cell.border = thick_border

            cell = ws_summary.cell(row=idx, column=3, value=row[1])
            cell.fill = fill_color
            cell.alignment = center_alignment
            if cell.value is not None:
                cell.border = thick_border

            cell = ws_summary.cell(row=idx, column=4, value=row[2])
            cell.fill = fill_color
            cell.alignment = center_alignment
            if cell.value is not None:
                cell.border = thick_border

            cell = ws_summary.cell(row=idx, column=6, value=row[3])
            cell.fill = fill_color
            cell.alignment = center_alignment
            if cell.value is not None:
                cell.border = thick_border

            cell = ws_summary.cell(row=idx, column=7, value=row[4])
            cell.fill = fill_color
            cell.alignment = center_alignment
            if cell.value is not None:
                cell.border = thick_border

        # Optionally adjust column widths for neat formatting
        for col in ['A','B','C','D','E','F','G','H','I']:
            ws_summary.column_dimensions[col].width = 20

        # Save the updated workbook into a BytesIO stream for download
        output = BytesIO()
        wb.save(output)
        processed_file = output.getvalue()

        st.download_button(
            label="Download Updated Excel File",
            data=processed_file,
            file_name="updated_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )