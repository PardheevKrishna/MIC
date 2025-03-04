import streamlit as st
import pandas as pd
import datetime
import re
from dateutil.relativedelta import relativedelta
from io import BytesIO

# -----------------------------------------------------------
# 1. PAGE CONFIG & CUSTOM STYLES
# -----------------------------------------------------------
st.set_page_config(
    page_title="BDCOMM FRY14M Field Analysis",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Inject some CSS to beautify the UI
custom_css = """
<style>
/* Gradient background for the main page */
[data-testid="stAppViewContainer"] {
    background: linear-gradient(135deg, #f0f7ff 0%, #dbefff 100%);
    color: #2F4F4F;
}

/* Center the main title and make it bigger */
h1 {
    text-align: center;
    font-family: 'Helvetica Neue', sans-serif;
    font-weight: 700;
    color: #2F4F4F;
}

/* Card-like box for results */
div.stContainer {
    background-color: rgba(255, 255, 255, 0.8);
    padding: 2rem;
    border-radius: 10px;
}

/* Buttons and inputs */
section.main > div:nth-child(1) {
    border-radius: 10px;
}

/* Style the table headers */
thead tr th {
    background-color: #4682B4 !important;
    color: white !important;
    font-weight: bold !important;
}

/* Thicker border for table cells */
tbody tr td, thead tr th {
    border: 1px solid #999 !important;
}

/* Subtle row hover effect */
tbody tr:hover {
    background-color: #FAFAD2 !important;
}
</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# -----------------------------------------------------------
# 2. SIDEBAR INPUTS
# -----------------------------------------------------------
st.sidebar.title("Settings")

# File uploader
uploaded_file = st.sidebar.file_uploader("Upload an Excel file (.xlsx)", type=["xlsx"])

# Date picker for date1 (default Jan 1, 2025)
selected_date = st.sidebar.date_input("Select Date for Date1", datetime.date(2025, 1, 1))
date1 = datetime.datetime.combine(selected_date, datetime.datetime.min.time())
date2 = date1 - relativedelta(months=1)  # 1 month before

# Action button
generate_button = st.sidebar.button("Generate Summary")

# -----------------------------------------------------------
# 3. MAIN TITLE
# -----------------------------------------------------------
st.title("BDCOMM FRY14M Field Analysis Summary")

st.markdown("""
**Instructions**:
1. Use the sidebar to upload your Excel file.
2. Pick the reference date (Date1). We'll automatically set Date2 to one month prior.
3. Click **Generate Summary** to view results.
""")

# -----------------------------------------------------------
# 4. PROCESS & DISPLAY RESULTS
# -----------------------------------------------------------
if generate_button:
    if not uploaded_file:
        st.error("Please upload an Excel file first.")
    else:
        try:
            # Read the "Data" sheet
            df_data = pd.read_excel(uploaded_file, sheet_name="Data")
        except Exception as e:
            st.error(f"Error reading 'Data' sheet: {e}")
            st.stop()

        # Convert the filemonth_dt to datetime
        try:
            df_data["filemonth_dt"] = pd.to_datetime(df_data["filemonth_dt"])
        except:
            st.warning("Could not parse 'filemonth_dt' properly. Make sure it's a valid date column.")
        
        # Ensure required columns exist
        required_cols = {"analysis_type", "filemonth_dt", "field_name", "value_label", "value_records"}
        if not required_cols.issubset(df_data.columns):
            st.error(f"Your 'Data' sheet must contain at least these columns: {required_cols}")
            st.stop()
        
        # Get unique, sorted field names
        fields = sorted(df_data["field_name"].unique())

        # Define phrases (with escaped parentheses)
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

        # Summarize data for each field
        summary_list = []
        for field in fields:
            # Missing values (analysis_type='value_dist', value_label contains 'Missing')
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

            # Month-to-month differences (analysis_type='pop_comp', value_label has the phrases)
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

            summary_list.append([field, missing_sum_d1, missing_sum_d2, m2m_sum_d1, m2m_sum_d2])

        # Convert summary list to a DataFrame
        summary_df = pd.DataFrame(
            summary_list,
            columns=[
                "Field Name",
                f"Missing Values ({date1.strftime('%Y-%m-%d')})",
                f"Missing Values ({date2.strftime('%Y-%m-%d')})",
                f"M2M Diff ({date1.strftime('%Y-%m-%d')})",
                f"M2M Diff ({date2.strftime('%Y-%m-%d')})"
            ]
        )

        # Display results
        st.subheader("Results")
        st.dataframe(summary_df, use_container_width=True)

        # Provide a quick way to download the summary as CSV
        csv_buffer = BytesIO()
        summary_df.to_csv(csv_buffer, index=False)
        st.download_button(
            label="Download Summary as CSV",
            data=csv_buffer.getvalue(),
            file_name="summary.csv",
            mime="text/csv"
        )

        # Celebration if no error
        st.balloons()