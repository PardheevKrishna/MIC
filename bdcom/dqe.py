import streamlit as st
st.set_page_config(page_title="DQE Analysis", layout="wide", initial_sidebar_state="expanded")

import pandas as pd
import os
import datetime
import re
from dateutil.relativedelta import relativedelta
from io import BytesIO
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode

#############################################
# Helper Functions
#############################################

def compute_grid_height(df, row_height=40, header_height=80):
    n = len(df)
    return header_height + (min(n, 30) * row_height)

def normalize_columns(df, mapping={"edit_nbr": "edit_nbr", "edit_error_cnt": "edit_error_cnt", "edit_threshold_cnt": "edit_threshold_cnt"}):
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
    # Read the DATA_INPUT sheet.
    df = pd.read_excel(excel_path, sheet_name="DATA_INPUT")
    df = normalize_columns(df)
    # Convert filemonth_dt to datetime and create a Year-Month string column.
    df["filemonth_dt"] = pd.to_datetime(df["filemonth_dt"])
    df["filemonth_str"] = df["filemonth_dt"].apply(format_date_to_ym)
    return df

def load_dqe_thresholds(csv_path):
    df = pd.read_csv(csv_path)
    df.columns = [str(col).strip() for col in df.columns]
    return df

#############################################
# DQE Analysis Processing
#############################################

def process_dqe_analysis(data_input_df, thresholds_df, date1):
    """
    For each row in thresholds_df, for each month in the range [Date1, Date1-4 months],
    sum edit_error_cnt and edit_threshold_cnt (grouped by filemonth_str and edit_nbr),
    compute Error%, and assign Status ("Pass" if Error% <= Threshold, else "Fail").
    An empty "Error Comments" column is added for each month.
    """
    # Define 5-month range: Date1 and the previous 4 months.
    months = [date1 - relativedelta(months=i) for i in range(5)]
    month_strs = [format_date_to_ym(m) for m in months]

    # Group the data input by filemonth_str and edit_nbr.
    grp = data_input_df.groupby(["filemonth_str", "edit_nbr"]).agg({
        "edit_error_cnt": "sum",
        "edit_threshold_cnt": "sum"
    }).reset_index()

    results = []
    for idx, thresh in thresholds_df.iterrows():
        # Use CSV column "Edit Nbr" to match data_input_df["edit_nbr"]
        edit_nbr_csv = thresh["Edit Nbr"]
        out_row = thresh.to_dict()  # include all CSV columns
        for m_str in month_strs:
            filt = grp[(grp["filemonth_str"] == m_str) & (grp["edit_nbr"] == edit_nbr_csv)]
            error_cnt = filt["edit_error_cnt"].sum() if not filt.empty else 0
            threshold_cnt = filt["edit_threshold_cnt"].sum() if not filt.empty else 0
            error_pct = (error_cnt / threshold_cnt * 100) if threshold_cnt > 0 else 0
            try:
                threshold_val = float(thresh["Threshold"])
            except:
                threshold_val = 0
            status = "Pass" if error_pct <= threshold_val else "Fail"
            out_row[f"{m_str} Errors"] = error_cnt
            out_row[f"{m_str} Error%"] = round(error_pct, 2)
            out_row[f"{m_str} Status"] = status
            out_row[f"{m_str} Error Comments"] = ""  # initially empty
        results.append(out_row)
    result_df = pd.DataFrame(results)
    # Replace any "nan" strings with empty strings.
    result_df = result_df.replace("nan", "")
    return result_df

#############################################
# Main Streamlit App for DQE Analysis
#############################################

def main():
    st.title("DQE Analysis Report")

    # Sidebar: Folder and file selection.
    folder = st.sidebar.selectbox("Select Folder", ["BDCOM", "WFHMSA", "BCards"])
    folder_path = os.path.join(os.getcwd(), folder)
    st.sidebar.write(f"Folder path: {folder_path}")
    if not os.path.exists(folder_path):
        st.sidebar.error(f"Folder '{folder}' not found.")
        return

    excel_files = [f for f in os.listdir(folder_path) if f.lower().endswith(('.xlsx','.xlsb'))]
    if not excel_files:
        st.sidebar.error(f"No Excel files found in folder '{folder}'.")
        return
    selected_excel = st.sidebar.selectbox("Select DATA_INPUT Excel File", excel_files)
    input_excel_path = os.path.join(folder_path, selected_excel)

    csv_path = os.path.join(folder_path, "dqe_thresholds.csv")
    if not os.path.exists(csv_path):
        st.sidebar.error(f"CSV file 'dqe_thresholds.csv' not found in folder '{folder}'.")
        return

    analysis_date = st.sidebar.date_input("Select Analysis Date (Date1)", datetime.date(2025, 1, 1))
    analysis_date_dt = datetime.datetime.combine(analysis_date, datetime.datetime.min.time())

    # Load DATA_INPUT and thresholds data.
    data_input_df = load_data_input(input_excel_path)
    thresholds_df = load_dqe_thresholds(csv_path)

    # Process the DQE analysis.
    dqe_result = process_dqe_analysis(data_input_df, thresholds_df, analysis_date_dt)

    # Replace any "nan" strings with empty strings.
    dqe_result = dqe_result.replace("nan", "")

    # Configure AgGrid to display the DQE analysis with filtering enabled on every column.
    gb = GridOptionsBuilder.from_dataframe(dqe_result)
    gb.configure_default_column(filter=True, editable=True, cellStyle={'white-space': 'normal', 'line-height': '1.2em', 'width': 150})
    # Make each "Error Comments" column editable.
    for col in dqe_result.columns:
        if "Error Comments" in col:
            gb.configure_column(col, editable=True, width=220)
    # Build grid options.
    grid_opts = gb.build()

    # Display the grid.
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

    # Provide a download button for the result as a new Excel file.
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