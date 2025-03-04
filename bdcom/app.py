import streamlit as st
import pandas as pd
import os
import datetime
import re
from io import BytesIO
from dateutil.relativedelta import relativedelta

def main():
    # -----------------------------------------------------------
    # 1. PAGE CONFIG
    # -----------------------------------------------------------
    st.set_page_config(
        page_title="Official FRY14M Field Analysis Summary",
        layout="centered",  # or "wide" if you prefer
        initial_sidebar_state="expanded"
    )

    # -----------------------------------------------------------
    # 2. SIDEBAR - FILE SELECTION AND DATE
    # -----------------------------------------------------------
    st.sidebar.title("File & Date Selection")
    
    # Gather .xlsx files in current directory
    xlsx_files = [f for f in os.listdir('.') if f.endswith('.xlsx')]
    
    if not xlsx_files:
        st.sidebar.error("No Excel files (.xlsx) found in the current folder.")
        return
    
    selected_file = st.sidebar.selectbox("Select an Excel File", xlsx_files)
    
    # Date input for Date1 (default: Jan 1, 2025)
    selected_date = st.sidebar.date_input(
        "Select Date for Date1",
        datetime.date(2025, 1, 1)
    )
    # Convert to datetime
    date1 = datetime.datetime.combine(selected_date, datetime.datetime.min.time())
    # Date2 is one month before Date1
    date2 = date1 - relativedelta(months=1)

    # Generate button
    if st.sidebar.button("Generate Summary"):
        generate_summary(selected_file, date1, date2)

def generate_summary(file_path, date1, date2):
    """
    Reads the Excel file's 'Data' sheet, computes the summary for Missing Values and 
    Month-to-Month Value Differences, and displays the results in the Streamlit app.
    """
    st.title("BDCOMM FRY14M Field Analysis Summary")

    # -----------------------------------------------------------
    # 3. READ DATA SHEET
    # -----------------------------------------------------------
    try:
        df_data = pd.read_excel(file_path, sheet_name="Data")
    except Exception as e:
        st.error(f"Could not read the 'Data' sheet from {file_path}. Error: {e}")
        return

    # Ensure required columns exist
    required_cols = {"analysis_type", "filemonth_dt", "field_name", "value_label", "value_records"}
    if not required_cols.issubset(df_data.columns):
        st.error(
            f"The 'Data' sheet must contain at least these columns: {required_cols}.\n"
            f"Columns found: {list(df_data.columns)}"
        )
        return

    # Convert filemonth_dt to datetime
    try:
        df_data["filemonth_dt"] = pd.to_datetime(df_data["filemonth_dt"])
    except:
        st.warning("Could not parse 'filemonth_dt' properly. Ensure it's a valid date column.")

    # -----------------------------------------------------------
    # 4. PREPARE FOR CALCULATIONS
    # -----------------------------------------------------------
    # Sort field names
    fields = sorted(df_data["field_name"].unique())

    # Phrases (with escaped parentheses) for pop_comp
    phrases = [
        "1\\)   F6CF Loan - Both Pop, Diff Values",
        "2\\)   CF Loan - Prior Null, Current Pop",
        "3\\)   CF Loan - Prior Pop, Current Null"
    ]

    def contains_phrase(text, patterns):
        for pat in patterns:
            if re.search(pat, text):
                return True
        return False

    # -----------------------------------------------------------
    # 5. COMPUTE SUMMARY
    # -----------------------------------------------------------
    summary_data = []
    for field in fields:
        # Missing values (analysis_type = 'value_dist', value_label contains 'Missing')
        mask_missing_date1 = (
            (df_data['analysis_type'] == 'value_dist') &
            (df_data['field_name'] == field) &
            (df_data['filemonth_dt'] == date1) &
            (df_data['value_label'].str.contains("Missing", case=False, na=False))
        )
        missing_sum_d1 = df_data.loc[mask_missing_date1, 'value_records'].sum()

        mask_missing_date2 = (
            (df_data['analysis_type'] == 'value_dist') &
            (df_data['field_name'] == field) &
            (df_data['filemonth_dt'] == date2) &
            (df_data['value_label'].str.contains("Missing", case=False, na=False))
        )
        missing_sum_d2 = df_data.loc[mask_missing_date2, 'value_records'].sum()

        # Month-to-month differences (analysis_type='pop_comp', value_label matches any phrase)
        mask_m2m_date1 = (
            (df_data['analysis_type'] == 'pop_comp') &
            (df_data['field_name'] == field) &
            (df_data['filemonth_dt'] == date1) &
            (df_data['value_label'].apply(lambda x: contains_phrase(x, phrases)))
        )
        m2m_sum_d1 = df_data.loc[mask_m2m_date1, 'value_records'].sum()

        mask_m2m_date2 = (
            (df_data['analysis_type'] == 'pop_comp') &
            (df_data['field_name'] == field) &
            (df_data['filemonth_dt'] == date2) &
            (df_data['value_label'].apply(lambda x: contains_phrase(x, phrases)))
        )
        m2m_sum_d2 = df_data.loc[mask_m2m_date2, 'value_records'].sum()

        summary_data.append([
            field,
            missing_sum_d1,
            missing_sum_d2,
            m2m_sum_d1,
            m2m_sum_d2
        ])

    # Create a DataFrame for easy display
    summary_df = pd.DataFrame(
        summary_data,
        columns=[
            "Field Name",
            f"Missing Values ({date1.strftime('%Y-%m-%d')})",
            f"Missing Values ({date2.strftime('%Y-%m-%d')})",
            f"Month-to-Month Diff ({date1.strftime('%Y-%m-%d')})",
            f"Month-to-Month Diff ({date2.strftime('%Y-%m-%d')})"
        ]
    )

    # -----------------------------------------------------------
    # 6. DISPLAY RESULTS
    # -----------------------------------------------------------
    st.write(f"**Selected File:** {file_path}")
    st.write(f"**Date1:** {date1.strftime('%Y-%m-%d')} | **Date2:** {date2.strftime('%Y-%m-%d')}")

    st.subheader("Summary Results")
    st.dataframe(summary_df, use_container_width=True)

    # -----------------------------------------------------------
    # 7. DOWNLOAD BUTTON
    # -----------------------------------------------------------
    csv_buffer = BytesIO()
    summary_df.to_csv(csv_buffer, index=False)
    st.download_button(
        label="Download Summary as CSV",
        data=csv_buffer.getvalue(),
        file_name="summary_report.csv",
        mime="text/csv"
    )

if __name__ == "__main__":
    main()