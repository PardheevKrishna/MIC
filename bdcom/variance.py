import streamlit as st
import pandas as pd
import os
from io import BytesIO

st.set_page_config(page_title="Variance Analysis", layout="wide", initial_sidebar_state="expanded")

#############################################
# Helper Functions
#############################################

def compute_grid_height(df, row_height=40, header_height=80):
    n = len(df)
    return header_height + (min(n, 30) * row_height)

def load_variance_data(excel_path):
    # Load both 'OUTPUT' and 'Variance_Analysis' sheets from the Excel file
    output_df = pd.read_excel(excel_path, sheet_name="OUTPUT")
    variance_df = pd.read_excel(excel_path, sheet_name="Variance_Analysis")
    
    # Clean column names to remove leading/trailing spaces
    output_df.columns = [col.strip() for col in output_df.columns]
    variance_df.columns = [col.strip() for col in variance_df.columns]
    
    return output_df, variance_df

def calculate_variances(output_df, variance_df):
    # Add "Current Value" column by summing the values from 'C' column in OUTPUT for matching 'Field Name'
    variance_df["Current Value"] = variance_df["Field Name"].apply(
        lambda x: output_df[output_df["Field Name"] == x]["C"].sum() if x in output_df["Field Name"].values else 0
    )
    
    # Add "Prior Value" column by summing the values from 'F' column in OUTPUT for matching 'Field Name'
    variance_df["Prior Value"] = variance_df["Field Name"].apply(
        lambda x: output_df[output_df["Field Name"] == x]["F"].sum() if x in output_df["Field Name"].values else 0
    )
    
    # Add "$Variance" column (Current Value - Prior Value)
    variance_df["$Variance"] = variance_df["Current Value"] - variance_df["Prior Value"]
    
    # Add "%Variance" column (Variance / Prior Value * 100), handle division by zero
    variance_df["%Variance"] = variance_df.apply(
        lambda row: (row["$Variance"] / row["Prior Value"] * 100) if row["Prior Value"] != 0 else 0,
        axis=1
    )
    
    # Return the final DataFrame with the additional columns
    return variance_df

def display_variance_analysis(df):
    # Display the grid with the added columns and allow editing
    import st_aggrid
    from st_aggrid import GridOptionsBuilder, GridUpdateMode, DataReturnMode

    gb = GridOptionsBuilder.from_dataframe(df)
    gb.configure_default_column(filter=True, editable=True, cellStyle={'white-space': 'normal', 'line-height': '1.2em', 'width': 150})
    # Make the "Error Comments" column editable for the user to write comments.
    for col in df.columns:
        if "Error Comments" in col:
            gb.configure_column(col, editable=True, width=220)
    
    grid_options = gb.build()
    
    st.subheader("Variance Analysis Table")
    AgGrid(
        df,
        gridOptions=grid_options,
        update_mode=GridUpdateMode.VALUE_CHANGED,
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        key="variance_grid",
        height=compute_grid_height(df, 40, 80),
        use_container_width=True
    )

    # Provide the download button for the updated DQE analysis as Excel
    towrite = BytesIO()
    with pd.ExcelWriter(towrite, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Variance Analysis")
    towrite.seek(0)
    st.download_button(
        label="Download Variance Analysis Report",
        data=towrite,
        file_name="Variance_Analysis_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

#############################################
# Main Streamlit App
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
    selected_excel = st.sidebar.selectbox("Select Excel File", excel_files)
    input_excel_path = os.path.join(folder_path, selected_excel)

    # Load the data from both the 'OUTPUT' and 'Variance_Analysis' sheets
    output_df, variance_df = load_variance_data(input_excel_path)

    # Perform the calculations for the variance analysis
    result_df = calculate_variances(output_df, variance_df)

    # Display the result in the AgGrid table
    display_variance_analysis(result_df)

if __name__ == "__main__":
    main()