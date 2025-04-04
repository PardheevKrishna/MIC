import streamlit as st
st.set_page_config(page_title="Variance Analysis", layout="wide", initial_sidebar_state="expanded")

import pandas as pd
import os
import datetime
import re
from io import BytesIO
from dateutil.relativedelta import relativedelta
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode
from openpyxl import load_workbook
from openpyxl.comments import Comment
import numpy as np

#############################################
# Helper Functions
#############################################

def compute_grid_height(df, row_height=40, header_height=80):
    n = len(df)
    return header_height + (min(n, 30) * row_height)

def normalize_columns(df):
    # Strip whitespace from column names.
    df.columns = [str(col).strip() for col in df.columns]
    return df

#############################################
# Data Loading Functions
#############################################

def load_output_data(excel_path):
    """
    Load the OUTPUT sheet from the Excel file.
    Assumes that the header row is the first row.
    We will use column indexes:
      - Column B (index 1) contains the field names (for Current Value search)
      - Column C (index 2) contains the numeric current values.
      - Column E (index 4) contains the field names (for Prior Value search)
      - Column F (index 5) contains the numeric prior values.
    """
    # Read the OUTPUT sheet (default header row 0)
    output_df = pd.read_excel(excel_path, sheet_name="OUTPUT", header=0)
    output_df = normalize_columns(output_df)
    return output_df

def load_variance_analysis_sheet(excel_path):
    """
    Load the Variance_Analysis sheet.
    The header row is in the 8th row (i.e. header=7 in 0-index).
    This sheet is expected to have columns such as:
      "14M file", "Field No.", "Field Name", "$Tolerance", "%Tolerance", etc.
    """
    var_df = pd.read_excel(excel_path, sheet_name="Variance_Analysis", header=7)
    var_df = normalize_columns(var_df)
    return var_df

#############################################
# Variance Analysis Processing
#############################################

def process_variance_analysis(output_df, var_df):
    """
    For each row in the Variance_Analysis DataFrame, compute:
      - "Current Value": Sum of values from OUTPUT sheet's column C (index 2)
         where OUTPUT sheet's column B (index 1) equals the "Field Name".
      - "Prior Value": Sum of values from OUTPUT sheet's column F (index 5)
         where OUTPUT sheet's column E (index 4) equals the "Field Name".
      - "$Variance" = Current Value - Prior Value.
      - "%Variance" = ($Variance / Prior Value) * 100 (0 if Prior Value is 0).
    The original columns from var_df are preserved (including "$Tolerance" and "%Tolerance").
    """
    # Ensure OUTPUT sheet is read as a DataFrame.
    # Use .iloc to reference columns by index.
    # For clarity, we assume:
    #   - output_df.iloc[:,1] is Field Name for Current Value search.
    #   - output_df.iloc[:,2] contains the current numeric values.
    #   - output_df.iloc[:,4] is Field Name for Prior Value search.
    #   - output_df.iloc[:,5] contains the prior numeric values.
    current_vals = []
    prior_vals = []
    
    # Loop through each row in var_df.
    for idx, row in var_df.iterrows():
        field = row["Field Name"]
        # Current Value: search in OUTPUT sheet column B (index 1) and sum corresponding column C (index 2).
        curr_val = output_df[output_df.iloc[:, 1] == field].iloc[:, 2].sum() if field in output_df.iloc[:,1].values else 0
        # Prior Value: search in OUTPUT sheet column E (index 4) and sum corresponding column F (index 5).
        prior_val = output_df[output_df.iloc[:, 4] == field].iloc[:, 5].sum() if field in output_df.iloc[:,4].values else 0
        current_vals.append(curr_val)
        prior_vals.append(prior_val)
    
    var_df["Current Value"] = current_vals
    var_df["Prior Value"] = prior_vals
    var_df["$Variance"] = var_df["Current Value"] - var_df["Prior Value"]
    var_df["%Variance"] = var_df.apply(lambda r: (r["$Variance"] / r["Prior Value"] * 100) if r["Prior Value"] != 0 else 0, axis=1)
    
    # Replace any "nan" strings with empty strings.
    var_df = var_df.replace("nan", "")
    return var_df

#############################################
# Main Streamlit App
#############################################

def main():
    st.title("Variance Analysis Report")

    # Sidebar: Select folder and file.
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
    selected_file = st.sidebar.selectbox("Select Excel File", excel_files)
    input_excel_path = os.path.join(folder_path, selected_file)

    # Load the two sheets: OUTPUT and Variance_Analysis.
    output_df = load_output_data(input_excel_path)
    var_df = load_variance_analysis_sheet(input_excel_path)

    st.write("### OUTPUT Sheet (first 5 rows)")
    st.dataframe(output_df.head())
    st.write("### Variance Analysis Sheet (first 5 rows)")
    st.dataframe(var_df.head())

    # Process the variance analysis.
    result_df = process_variance_analysis(output_df, var_df)

    # Replace any literal "nan" with empty strings.
    result_df = result_df.replace("nan", "")

    # Use AgGrid to display the result.
    from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode
    gb = GridOptionsBuilder.from_dataframe(result_df)
    # Enable filtering on every column.
    gb.configure_default_column(filter=True, editable=True, cellStyle={'white-space': 'normal', 'line-height': '1.2em', 'width': 150})
    # Pin the "Field Name" column and "Field No." if available.
    gb.configure_column("Field Name", pinned="left")
    if "Field No." in result_df.columns:
        gb.configure_column("Field No.", pinned="left")
    # Make the "Error Comments" column editable if you want users to add comments.
    # (If you want a new column for comments, add it here.)
    if "Error Comments" not in result_df.columns:
        result_df["Error Comments"] = ""
    gb.configure_column("Error Comments", editable=True, width=220)

    grid_opts = gb.build()

    AgGrid(
        result_df,
        gridOptions=grid_opts,
        update_mode=GridUpdateMode.VALUE_CHANGED,
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        key="variance_grid",
        height=compute_grid_height(result_df, 40, 80),
        use_container_width=True
    )

    # Download button for the result as a new Excel file.
    towrite = BytesIO()
    with pd.ExcelWriter(towrite, engine="openpyxl") as writer:
        result_df.to_excel(writer, index=False, sheet_name="Variance Analysis")
    towrite.seek(0)
    st.download_button(
        label="Download Variance Analysis Report as Excel",
        data=towrite,
        file_name="Variance_Analysis_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    main()