import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import os

# ---------- CONFIGURATION ----------
FILE_PATH = "input.xlsx"  # Change this to your actual Excel file path

# -------------- CUSTOM CSS & TITLE --------------
st.markdown("""
<style>
.mySkipButton { visibility: hidden; }
</style>
""", unsafe_allow_html=True)
st.title("Team Report Dashboard")

# ===================== DATA LOADING =====================
def load_data(file_path):
    """
    In your real code, replace this function with your process_excel_file logic.
    Here we load two DataFrames: working_details and violations_df.
    """
    if os.path.exists(file_path):
        try:
            # Replace the following with your actual Excel reading logic.
            working_details = pd.read_excel(file_path, sheet_name=0)
            violations_df = pd.read_excel(file_path, sheet_name=1)
        except Exception as e:
            st.error(f"Error reading Excel file: {e}")
            st.stop()
    else:
        # Create dummy data for demonstration
        data = {
            "Employee": ["Alice", "Bob", "Charlie"],
            "UniqueID": ["Alice_2", "Bob_3", "Charlie_4"],
            "Start Date": ["01-01-2023", "01-02-2023", "01-03-2023"],
            "Completion Date": ["01-10-2023", "01-11-2023", "01-12-2023"],
            "RowNumber": [2, 3, 4]
        }
        working_details = pd.DataFrame(data)
        violations_df = pd.DataFrame(data)
    return working_details, violations_df

working_details, violations_df = load_data(FILE_PATH)
# Store the dataframes in session state if not already there
st.session_state.setdefault("working_details", working_details)
st.session_state.setdefault("violations_df", violations_df)

# ---------- USE st.radio TO SELECT A TAB (this preserves the active view) ----------
tab_option = st.radio("Select Report Tab", 
                       ["Team Monthly Summary", "Working Hours Summary", "Violations and Update"],
                       index=2)

# ---------- TEAM MONTHLY SUMMARY (Tab 1) ----------
if tab_option == "Team Monthly Summary":
    st.subheader("Team Monthly Summary")
    st.write("Team Monthly Summary code goes here.")

# ---------- WORKING HOURS SUMMARY (Tab 2) ----------
elif tab_option == "Working Hours Summary":
    st.subheader("Working Hours Summary")
    st.write("Working Hours Summary code goes here.")

# ---------- VIOLATIONS & UPDATE (Tab 3) ----------
elif tab_option == "Violations and Update":
    st.subheader("Violations and Update")
    # Display the violations dataframe
    st.dataframe(st.session_state["violations_df"])

    # --- Step 1: Row Selection (in a form) ---
    with st.form("row_selection_form"):
        all_ids = st.session_state["violations_df"]["UniqueID"].tolist()
        select_all = st.checkbox("Select All Rows", key="select_all_rows")
        if select_all:
            selected_ids = all_ids
        else:
            selected_ids = st.multiselect("Select UniqueIDs", options=all_ids, key="selected_ids")
        load_submitted = st.form_submit_button("Load Editing Forms")
    
    if load_submitted:
        if not selected_ids:
            st.error("No rows selected for update.")
        else:
            st.session_state["selected_rows"] = selected_ids
            st.success(f"Selected rows: {selected_ids}")
    
    # --- Step 2: Deferred Editing Forms ---
    if "selected_rows" in st.session_state:
        st.markdown("### Edit Each Selected Row")
        updated_data = {}  # Dictionary to collect new values
        # List of fields to edit; adjust as needed.
        edit_fields = ["Start Date", "Completion Date"]
        # Build a lookup from UniqueID to corresponding row in working_details
        wd = st.session_state["working_details"]
        working_details_dict = {row["UniqueID"]: row for _, row in wd.iterrows()}
        
        for uid in st.session_state["selected_rows"]:
            if uid not in working_details_dict:
                st.warning(f"No data found for {uid}")
                continue
            row = working_details_dict[uid]
            with st.expander(f"Edit row {uid}", expanded=True):
                # Use text inputs with unique keys
                new_start = st.text_input("Start Date", value=str(row["Start Date"]), key=f"{uid}_start")
                new_comp = st.text_input("Completion Date", value=str(row.get("Completion Date", "")), key=f"{uid}_comp")
                # Collect updated values
                updated_data[uid] = {
                    "Employee": row["Employee"],
                    "RowNumber": row.get("RowNumber", 2),  # Adjust row number if needed
                    "Start Date": new_start,
                    "Completion Date": new_comp
                }
                # If you have additional fields, add more inputs here.
        
        # --- Step 3: Update Excel via Pure Python Function ---
        if st.button("Update Excel", key="update_excel"):
            def update_excel_file_py(updated_data):
                try:
                    wb = load_workbook(FILE_PATH)
                except Exception as e:
                    st.error(f"Error opening workbook: {e}")
                    return False
                for uid, row_vals in updated_data.items():
                    sheet_name = row_vals["Employee"]
                    if sheet_name not in wb.sheetnames:
                        st.warning(f"Sheet {sheet_name} not found.")
                        continue
                    ws = wb[sheet_name]
                    # Assumes the first row in the sheet contains headers
                    headers = {cell.value: cell.column for cell in ws[1]}
                    r_num = row_vals["RowNumber"]
                    if "Start Date" in headers and row_vals["Start Date"]:
                        ws.cell(row=r_num, column=headers["Start Date"], value=row_vals["Start Date"])
                    if "Completion Date" in headers and row_vals["Completion Date"]:
                        ws.cell(row=r_num, column=headers["Completion Date"], value=row_vals["Completion Date"])
                    # Extend this logic for additional fields if necessary.
                try:
                    wb.save(FILE_PATH)
                    return True
                except Exception as e:
                    st.error(f"Error saving workbook: {e}")
                    return False

            if update_excel_file_py(updated_data):
                st.success("Excel file updated successfully (via pure Python function).")
                # Optionally update the session_state working_details so that changes appear immediately.
                for uid, row_vals in updated_data.items():
                    mask = st.session_state["working_details"]["UniqueID"] == uid
                    if mask.any():
                        idx = st.session_state["working_details"].index[mask][0]
                        for col, new_val in row_vals.items():
                            st.session_state["working_details"].at[idx, col] = new_val
                st.success("Session data updated. Changes now appear in your dashboard.")