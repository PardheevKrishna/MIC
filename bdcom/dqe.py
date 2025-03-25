import streamlit as st
st.set_page_config(page_title="DQE Analysis", layout="wide", initial_sidebar_state="expanded")

import pandas as pd
import os
import datetime
import re
from dateutil.relativedelta import relativedelta
from io import BytesIO

#############################################
# Helper Functions
#############################################

def compute_grid_height(df, row_height=40, header_height=80):
    n = len(df)
    return header_height + (min(n, 30) * row_height)

def normalize_columns(df, mapping=None):
    if mapping is None:
        mapping = {"filemonth": "filemonth", "edit_nbr": "edit_nbr"}
    df.columns = [str(col).strip() for col in df.columns]
    for orig, new in mapping.items():
        for col in df.columns:
            if col.lower() == orig.lower() and col != new:
                df.rename(columns={col: new}, inplace=True)
    return df

def format_date_to_ym(date_obj):
    return date_obj.strftime("%Y-%m")

#############################################
# Data Loading Functions
#############################################

def load_data_input(excel_path):
    # Reads the DATA_INPUT sheet from the Excel file.
    df = pd.read_excel(excel_path, sheet_name="DATA_INPUT")
    df = normalize_columns(df, mapping={"filemonth": "filemonth", "edit_nbr": "edit_nbr",
                                          "edit_error_cnt": "edit_error_cnt", "edit_threshold_cnt": "edit_threshold_cnt"})
    # Convert filemonth column to datetime and add a Year-Month string column.
    df["filemonth_dt"] = pd.to_datetime(df["filemonth"])
    df["filemonth_str"] = df["filemonth_dt"].apply(format_date_to_ym)
    return df

def load_dqe_thresholds(csv_path):
    # Reads the thresholds CSV.
    df = pd.read_csv(csv_path)
    df.columns = [str(col).strip() for col in df.columns]
    return df

#############################################
# DQE Analysis Processing
#############################################

def process_dqe_analysis(data_input_df, thresholds_df, date1):
    """
    For each row in thresholds_df, for each month from date1 to date1-4 months,
    find matching rows in data_input_df (where edit_nbr matches CSV "Edit Nbr").
    For each month, sum up edit_error_cnt and edit_threshold_cnt, calculate Error%,
    assign Status ("Pass" if error percentage <= Threshold, else "Fail"),
    and create an empty "Error Comments" column for user input.
    """
    # Define 5-month range: Date1 and the previous 4 months.
    months = [date1 - relativedelta(months=i) for i in range(5)]
    month_strs = [format_date_to_ym(m) for m in months]

    # Group the data input by month and edit_nbr.
    grouped = data_input_df.groupby(["filemonth_str", "edit_nbr"]).agg({
        "edit_error_cnt": "sum",
        "edit_threshold_cnt": "sum"
    }).reset_index()

    results = []
    for idx, thresh in thresholds_df.iterrows():
        # Use the CSV column "Edit Nbr" to match data_input_df's "edit_nbr"
        edit_nbr_csv = thresh["Edit Nbr"]
        out_row = thresh.to_dict()  # include all CSV columns
        for m_str in month_strs:
            filt = grouped[(grouped["filemonth_str"] == m_str) & (grouped["edit_nbr"] == edit_nbr_csv)]
            error_cnt = filt["edit_error_cnt"].sum() if not filt.empty else 0
            threshold_cnt = filt["edit_threshold_cnt"].sum() if not filt.empty else 0
            error_pct = (error_cnt / threshold_cnt * 100) if threshold_cnt > 0 else 0
            try:
                threshold_val = float(thresh["Threshold"])
            except:
                threshold_val = 0
            status = "Pass" if error_pct <= threshold_val else "Fail"
            # Create four new columns per month.
            out_row[f"{m_str} Errors"] = error_cnt
            out_row[f"{m_str} Error%"] = round(error_pct, 2)
            out_row[f"{m_str} Status"] = status
            out_row[f"{m_str} Error Comments"] = ""  # empty, to be filled by user
        results.append(out_row)
    result_df = pd.DataFrame(results)
    return result_df

#############################################
# Main Streamlit App
#############################################

def main():
    st.title("DQE Analysis Report")

    # Sidebar: Select folder and file.
    folder = st.sidebar.selectbox("Select Folder", ["BDCOM", "WFHMSA", "BCards"])
    folder_path = os.path.join(os.getcwd(), folder)
    st.sidebar.write(f"Folder path: {folder_path}")
    if not os.path.exists(folder_path):
        st.sidebar.error(f"Folder '{folder}' not found.")
        return

    # Select the DATA_INPUT Excel file.
    excel_files = [f for f in os.listdir(folder_path) if f.lower().endswith(('.xlsx','.xlsb'))]
    if not excel_files:
        st.sidebar.error(f"No Excel files found in folder '{folder}'.")
        return
    selected_excel = st.sidebar.selectbox("Select DATA_INPUT Excel File", excel_files)
    input_excel_path = os.path.join(folder_path, selected_excel)

    # CSV file for thresholds (assumed to be in the same folder).
    csv_path = os.path.join(folder_path, "dqe_thresholds.csv")
    if not os.path.exists(csv_path):
        st.sidebar.error(f"CSV file 'dqe_thresholds.csv' not found in folder '{folder}'.")
        return

    # Date input for analysis (Date1).
    analysis_date = st.sidebar.date_input("Select Analysis Date (Date1)", datetime.date(2025, 1, 1))
    analysis_date_dt = datetime.datetime.combine(analysis_date, datetime.datetime.min.time())

    # Load data.
    data_input_df = load_data_input(input_excel_path)
    thresholds_df = load_dqe_thresholds(csv_path)

    st.write("### DATA_INPUT (first 5 rows)")
    st.dataframe(data_input_df.head())

    st.write("### dqe_thresholds.csv (first 5 rows)")
    st.dataframe(thresholds_df.head())

    # Process DQE analysis.
    dqe_result = process_dqe_analysis(data_input_df, thresholds_df, analysis_date_dt)
    # Replace any "nan" strings with empty strings.
    dqe_result = dqe_result.replace("nan", "")

    st.write("### DQE Analysis Result")
    st.dataframe(dqe_result)

    # Use AgGrid to allow editing of "Error Comments" columns.
    gb = GridOptionsBuilder.from_dataframe(dqe_result)
    gb.configure_default_column(editable=True, cellStyle={'white-space':'normal','line-height':'1.2em','width':150})
    # Ensure the "Error Comments" columns are editable.
    for col in dqe_result.columns:
        if "Error Comments" in col:
            gb.configure_column(col, editable=True, width=220)
        # Pin the original CSV columns if desired.
        if col in thresholds_df.columns:
            gb.configure_column(col, pinned="left")
    grid_opts = gb.build()

    grid_response = AgGrid(
        dqe_result,
        gridOptions=grid_opts,
        update_mode=GridUpdateMode.VALUE_CHANGED,
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        key="dqe_grid",
        height=compute_grid_height(dqe_result, 40, 80),
        use_container_width=True
    )
    updated_dqe = pd.DataFrame(grid_response["data"]).replace("nan", "", regex=True)

    st.write("### Updated DQE Analysis (after editing Error Comments)")
    st.dataframe(updated_dqe)

    # Provide a download button for the DQE Analysis report as a new Excel file.
    towrite = BytesIO()
    with pd.ExcelWriter(towrite, engine="openpyxl") as writer:
        updated_dqe.to_excel(writer, index=False, sheet_name="DQE Analysis")
    towrite.seek(0)
    st.download_button(
        label="Download DQE Analysis Report as Excel",
        data=towrite,
        file_name="DQE_Analysis_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    main()