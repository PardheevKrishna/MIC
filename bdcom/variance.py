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
    # Remove leading/trailing spaces from column names.
    df.columns = [str(col).strip() for col in df.columns]
    return df

def format_date_to_ym(date_obj):
    return date_obj.strftime("%Y-%m")

#############################################
# Data Loading Functions
#############################################

def load_output_data(excel_path):
    """
    Loads the OUTPUT sheet.
    Assumes:
      - Column B (index 1) contains field names for current values.
      - Column C (index 2) contains the current numeric values.
      - Column E (index 4) contains field names for prior values.
      - Column F (index 5) contains the prior numeric values.
    """
    output_df = pd.read_excel(excel_path, sheet_name="OUTPUT", header=0)
    output_df = normalize_columns(output_df)
    return output_df

def load_variance_analysis_sheet(excel_path):
    """
    Loads the Variance_Analysis sheet with header in row 8 (header=7).
    Expected columns include:
      "14M file", "Field No." (as given in the sheet), "Field Name", "$Tolerance", "%Tolerance", etc.
    We do not rename the "Field No." column so its values are preserved.
    """
    var_df = pd.read_excel(excel_path, sheet_name="Variance_Analysis", header=7)
    var_df = normalize_columns(var_df)
    return var_df

#############################################
# Variance Analysis Processing
#############################################

def process_variance_analysis(output_df, var_df):
    """
    For each row in the Variance_Analysis sheet, calculates:
      - "Current Value": Sum of values from OUTPUT sheet's column C
          where OUTPUT sheet's column B equals the "Field Name".
      - "Prior value": Sum of values from OUTPUT sheet's column F
          where OUTPUT sheet's column E equals the "Field Name".
      - "$Variance" = Current Value - Prior value.
      - "%Variance" = ($Variance / Prior value * 100) (0 if Prior value is 0).
    Ensures that editable columns "Comments" and "Detail File Link" exist.
    Reorders the columns into the following order:
      ["14M file", "Field No.", "Field Name", "Current Value", "Prior value",
       "$Variance", "%Variance", "$Tolerance", "%Tolerance", "Comments", "Detail File Link"]
    """
    current_values = []
    prior_values = []
    
    for idx, row in var_df.iterrows():
        field_name = row["Field Name"]
        # Current Value: search OUTPUT sheet column B (index 1) for field_name and sum corresponding column C (index 2).
        curr_val = 0
        if field_name in output_df.iloc[:, 1].values:
            curr_val = output_df[output_df.iloc[:, 1] == field_name].iloc[:, 2].sum()
        
        # Prior Value: search OUTPUT sheet column E (index 4) for field_name and sum corresponding column F (index 5).
        prior_val = 0
        if field_name in output_df.iloc[:, 4].values:
            prior_val = output_df[output_df.iloc[:, 4] == field_name].iloc[:, 5].sum()
        
        current_values.append(curr_val)
        prior_values.append(prior_val)
    
    var_df["Current Value"] = current_values
    var_df["Prior value"] = prior_values
    var_df["$Variance"] = var_df["Current Value"] - var_df["Prior value"]
    var_df["%Variance"] = var_df.apply(lambda r: (r["$Variance"] / r["Prior value"] * 100) if r["Prior value"] != 0 else 0, axis=1)
    
    # Ensure the editable columns exist.
    if "Comments" not in var_df.columns:
        var_df["Comments"] = ""
    if "Detail File Link" not in var_df.columns:
        var_df["Detail File Link"] = ""
    
    # Define the desired final column order.
    final_cols = ["14M file", "Field No.", "Field Name", "Current Value", "Prior value",
                  "$Variance", "%Variance", "$Tolerance", "%Tolerance", "Comments", "Detail File Link"]
    # Add missing columns if necessary.
    for col in final_cols:
        if col not in var_df.columns:
            var_df[col] = ""
    var_df = var_df[final_cols]
    return var_df

#############################################
# Main Streamlit App for Variance Analysis
#############################################

def main():
    st.title("Variance Analysis Report")

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
    selected_file = st.sidebar.selectbox("Select Excel File", excel_files)
    input_excel_path = os.path.join(folder_path, selected_file)

    # Load the OUTPUT sheet and Variance_Analysis sheet.
    output_df = load_output_data(input_excel_path)
    var_df = load_variance_analysis_sheet(input_excel_path)

    # Process the variance analysis.
    result_df = process_variance_analysis(output_df, var_df)
    result_df = result_df.replace("nan", "")  # Replace literal "nan" with empty strings

    # Configure AgGrid options with filtering enabled on every column.
    gb = GridOptionsBuilder.from_dataframe(result_df)
    gb.configure_default_column(filter=True, editable=True, cellStyle={'white-space':'normal', 'line-height':'1.2em', 'width':150})
    # Pin key columns.
    gb.configure_column("14M file", pinned="left")
    gb.configure_column("Field No.", pinned="left")
    gb.configure_column("Field Name", pinned="left")
    # Ensure that the editable columns "Comments" and "Detail File Link" are set.
    gb.configure_column("Comments", editable=True, width=220)
    gb.configure_column("Detail File Link", editable=True, width=220)
    grid_opts = gb.build()

    # Display the AgGrid table.
    grid_response = AgGrid(
        result_df,
        gridOptions=grid_opts,
        update_mode=GridUpdateMode.VALUE_CHANGED,
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        key="variance_grid",
        height=compute_grid_height(result_df, 40, 80),
        use_container_width=True
    )
    # Get updated data from AgGrid (which includes any edits in "Comments" and "Detail File Link").
    updated_df = pd.DataFrame(grid_response["data"]).replace("nan", "", regex=True)

    # Download: Create an Excel file with two sheets: "Variance Analysis" and "OUTPUT"
    towrite = BytesIO()
    with pd.ExcelWriter(towrite, engine="openpyxl") as writer:
        updated_df.to_excel(writer, index=False, sheet_name="Variance Analysis")
        output_df.to_excel(writer, index=False, sheet_name="OUTPUT")
    towrite.seek(0)
    st.download_button(
        label="Download Variance Analysis Report as Excel",
        data=towrite,
        file_name="Variance_Analysis_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    main()