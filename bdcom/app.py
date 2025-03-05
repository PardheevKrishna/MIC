import streamlit as st
import pandas as pd
import os
import datetime
import re
from io import BytesIO
from dateutil.relativedelta import relativedelta
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils.dataframe import dataframe_to_rows

# --------------------------------------------------------------------
# Function to build a detailed sheet (for a given analysis_type)
# --------------------------------------------------------------------
def generate_detailed_sheet(ws, analysis_type, df_data, period_list):
    """
    Writes a detailed monthly breakdown into the given worksheet.
    
    For the given analysis_type (either "value_dist" or "pop_comp"), the sheet
    groups by field_name. Under each field, it lists every value_label along with,
    for each month in period_list (from Date1 back 12 months), the sum of value_records
    and its percentage relative to the field’s total. Finally, a summary row ("Current period total")
    is added for each field.
    """
    # Set the sheet title (this will override the default name)
    title = "Value Distribution" if analysis_type == "value_dist" else "Population Comparison"
    ws.title = title

    # Build header row:
    header = ["Field Name", "Value Label"]
    for dt in period_list:
        month_label = dt.strftime('%Y-%m')
        header.append(f"Sum ({month_label})")
        header.append(f"Percentage ({month_label})")
    ws.append(header)
    # Bold the header row
    for cell in ws[1]:
        cell.font = Font(bold=True)

    # Filter the data for the current analysis_type
    df_subset = df_data[df_data['analysis_type'] == analysis_type].copy()
    # Create a new column 'period' as a year‐month string (e.g., "2025-01")
    df_subset['period'] = df_subset['filemonth_dt'].dt.to_period('M').astype(str)

    # Process each field name in sorted order
    field_names = sorted(df_subset['field_name'].unique())
    for field in field_names:
        # Write a row as a header for the field (merge the first two columns visually)
        num_cols = len(header)
        ws.append([f"Field Name: {field}"] + [""] * (num_cols - 1))
        
        # Subset for this field
        df_field = df_subset[df_subset['field_name'] == field]
        # Get all distinct value_labels for the field
        value_labels = sorted(df_field['value_label'].unique())
        for val in value_labels:
            row = ["", val]
            for dt in period_list:
                period_str = dt.strftime('%Y-%m')
                # Sum for this value_label in the given month
                df_temp = df_field[(df_field['value_label'] == val) & (df_field['period'] == period_str)]
                sum_val = df_temp['value_records'].sum()
                # Total for the entire field in this month
                total_field = df_field[df_field['period'] == period_str]['value_records'].sum()
                perc = (sum_val / total_field * 100) if total_field != 0 else 0
                row.append(sum_val)
                row.append(round(perc, 2))
            ws.append(row)
        # Now add a row for the "Current period total" for this field
        total_row = ["", "Current period total"]
        for dt in period_list:
            period_str = dt.strftime('%Y-%m')
            total_val = df_field[df_field['period'] == period_str]['value_records'].sum()
            total_row.append(total_val)
            total_row.append("")  # No percentage for the total row
        ws.append(total_row)
        # Add an empty row for spacing between fields
        ws.append([""] * len(header))

# --------------------------------------------------------------------
# Main Streamlit app
# --------------------------------------------------------------------
def main():
    st.set_page_config(
        page_title="Official FRY14M Field Analysis Report",
        layout="wide"
    )
    st.sidebar.title("File & Date Selection")

    # List all .xlsx files in the current directory for selection
    xlsx_files = [f for f in os.listdir('.') if f.endswith('.xlsx')]
    if not xlsx_files:
        st.sidebar.error("No Excel files (.xlsx) found in the folder.")
        st.stop()
    selected_file = st.sidebar.selectbox("Select an Excel file", xlsx_files)

    # Date picker for Date1 (default: January 1, 2025)
    selected_date = st.sidebar.date_input("Select Date for Date1", datetime.date(2025, 1, 1))
    date1 = datetime.datetime.combine(selected_date, datetime.datetime.min.time())
    
    # Create a list of months from Date1 back 12 months (chronological order)
    period_list = [date1 - relativedelta(months=i) for i in range(12)]
    period_list = sorted(period_list)

    if st.sidebar.button("Generate Report"):
        try:
            # Read the "Data" sheet from the selected Excel file
            df_data = pd.read_excel(selected_file, sheet_name="Data")
        except Exception as e:
            st.error(f"Error reading 'Data' sheet from {selected_file}: {e}")
            st.stop()

        # Ensure that filemonth_dt is parsed as datetime
        try:
            df_data['filemonth_dt'] = pd.to_datetime(df_data['filemonth_dt'])
        except Exception as e:
            st.error("Error converting 'filemonth_dt' column to datetime. Please check the date format.")
            st.stop()

        # ----------------------------
        # Build the Summary sheet
        # ----------------------------
        # For demonstration, create a simple summary that shows, for each field,
        # the total value_records for Date1 and Date2 (one month before Date1).
        date2 = date1 - relativedelta(months=1)
        fields = sorted(df_data['field_name'].unique())
        summary_list = []
        for field in fields:
            mask_d1 = (df_data['field_name'] == field) & (df_data['filemonth_dt'] == date1)
            mask_d2 = (df_data['field_name'] == field) & (df_data['filemonth_dt'] == date2)
            total_d1 = df_data.loc[mask_d1, 'value_records'].sum()
            total_d2 = df_data.loc[mask_d2, 'value_records'].sum()
            summary_list.append([field, total_d1, total_d2])
        summary_df = pd.DataFrame(
            summary_list, 
            columns=["Field Name", f"Total ({date1.strftime('%Y-%m-%d')})", f"Total ({date2.strftime('%Y-%m-%d')})"]
        )

        # ----------------------------
        # Create the Excel workbook with 3 sheets:
        # 1. Summary
        # 2. Value Distribution (for analysis_type = "value_dist")
        # 3. Population Comparison (for analysis_type = "pop_comp")
        # ----------------------------
        wb = Workbook()
        # --- Summary sheet ---
        ws_summary = wb.active
        ws_summary.title = "Summary"
        # Write the summary_df using openpyxl’s dataframe_to_rows helper
        for r_idx, row in enumerate(dataframe_to_rows(summary_df, index=False, header=True), start=1):
            ws_summary.append(row)
            if r_idx == 1:  # Bold header row
                for cell in ws_summary[r_idx]:
                    cell.font = Font(bold=True)
        
        # --- Value Distribution sheet ---
        ws_valdist = wb.create_sheet("Value Distribution")
        generate_detailed_sheet(ws_valdist, "value_dist", df_data, period_list)
        
        # --- Population Comparison sheet ---
        ws_popcomp = wb.create_sheet("Population Comparison")
        generate_detailed_sheet(ws_popcomp, "pop_comp", df_data, period_list)

        # Save the workbook to a BytesIO stream for download
        output = BytesIO()
        wb.save(output)
        processed_file = output.getvalue()

        st.success("Report generated successfully.")
        st.download_button(
            label="Download Report Excel File",
            data=processed_file,
            file_name="FRY14M_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Also show a preview of the Summary sheet as a table
        st.subheader("Summary Preview")
        st.dataframe(summary_df)

if __name__ == "__main__":
    main()