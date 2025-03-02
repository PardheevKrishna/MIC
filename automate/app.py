import streamlit as st
import pandas as pd
import json
from datetime import datetime, timedelta
from openpyxl import load_workbook
from io import BytesIO
import os

# ---------- CONFIGURATION ----------
FILE_PATH = "input.xlsx"        # Path to your main Excel file
TEMP_UPDATE_FILE = "temp_update.json"  # Temporary file to store update instructions

st.set_page_config(page_title="Team Report Dashboard", layout="wide")
st.title("Team Report Dashboard")

# ===================== PROCESS EXCEL FILE =====================
def process_excel_file(file_path):
    """
    Reads each employee sheet (employee names in the 'Home' sheet),
    and returns two DataFrames:
      - working_details: all rows with extra columns (Employee, RowNumber, Month, etc.)
      - violations_df: flagged violations
    """
    print("DEBUG: Entering process_excel_file...")
    # Example: define some allowed values and exceptions
    allowed_values = {
        "Functional Area (CRIT, CRIT - Data Management, CRIT - Data Governance, CRIT - Regulatory Reporting, CRIT - Portfolio Reporting, CRIT - Transformation)":
            ["CRIT", "CRIT - Data Management", "CRIT - Data Governance", "CRIT - Regulatory Reporting", "CRIT - Portfolio Reporting", "CRIT - Transformation"],
        # ... define the other categorical columns ...
    }
    start_date_exceptions = [
        "Internal meetings", "Internal Meetings", "internal meeting", "Interview",
        # etc...
    ]

    try:
        print("DEBUG: Reading 'Home' sheet to get employee names.")
        home_df = pd.read_excel(file_path, sheet_name="Home", header=None)
    except Exception as e:
        print(f"DEBUG: Error reading Home sheet: {e}")
        st.error(f"Error reading Home sheet: {e}")
        return None, None

    employee_names = home_df.iloc[2:, 5].dropna().astype(str).tolist()
    print("DEBUG: Employee names found:", employee_names)

    try:
        print("DEBUG: Loading entire Excel for sheet name checks.")
        xls = pd.ExcelFile(file_path)
        all_sheet_names = xls.sheet_names
        print("DEBUG: All sheet names in workbook:", all_sheet_names)
    except Exception as e:
        print(f"DEBUG: Error reading entire Excel file: {e}")
        st.error(f"Error reading Excel file: {e}")
        return None, None

    working_list = []
    viol_list = []
    project_month_info = {}

    for emp in employee_names:
        if emp not in all_sheet_names:
            print(f"DEBUG: Skipping employee '{emp}' - no matching sheet.")
            continue
        try:
            print(f"DEBUG: Reading sheet for employee: {emp}")
            df = pd.read_excel(file_path, sheet_name=emp)
        except Exception as e:
            print(f"DEBUG: Could not read sheet for {emp}: {e}")
            st.warning(f"Could not read sheet for {emp}: {e}")
            continue

        df.columns = [str(c).replace("\n", " ").strip() for c in df.columns]
        req_cols = ["Status Date (Every Friday)", "Main project", "Name of the Project", "Start Date", "Weekly Time Spent(Hrs)"]
        if not all(c in df.columns for c in req_cols):
            print(f"DEBUG: Skipping {emp}, missing required columns.")
            continue

        df["Employee"] = emp
        df["RowNumber"] = df.index + 2
        df["Status Date (Every Friday)"] = pd.to_datetime(
            df["Status Date (Every Friday)"], format="%m-%d-%Y", errors="coerce"
        )

        # ... check for allowed values, start date consistency, etc. ...
        # We'll skip the details for brevity but keep a debug line:
        print(f"DEBUG: Processed sheet '{emp}' with {len(df)} rows.")
        df["Month"] = df["Status Date (Every Friday)"].dt.to_period("M").astype(str)
        df["UniqueID"] = df["Employee"] + "_" + df["RowNumber"].astype(str)
        working_list.append(df)

    if working_list:
        working_details = pd.concat(working_list, ignore_index=True)
        print("DEBUG: Combined working_details shape:", working_details.shape)
    else:
        working_details = pd.DataFrame()
        print("DEBUG: No data found for working_details.")

    # Suppose we skip the actual violation building for brevity
    violations_df = pd.DataFrame(viol_list)
    print("DEBUG: Exiting process_excel_file.")
    return working_details, violations_df

# ---- LOAD DATA ----
print("DEBUG: Loading data via process_excel_file.")
working_details, violations_df = process_excel_file(FILE_PATH)
if working_details is None:
    print("DEBUG: working_details is None, stopping.")
    st.stop()
else:
    st.success("Data loaded successfully!")
    print("DEBUG: Data loaded successfully, shape:", working_details.shape)

# ========== TABS ==========
tab_option = st.radio("Select a Tab", [
    "Team Monthly Summary",
    "Working Hours Summary",
    "Violations",
    "Update Data"
])

if tab_option == "Team Monthly Summary":
    st.subheader("Team Monthly Summary")
    st.write("Placeholder for summary logic.")
    print("DEBUG: User is in 'Team Monthly Summary' tab.")

elif tab_option == "Working Hours Summary":
    st.subheader("Working Hours Summary")
    st.write("Placeholder for working hours logic.")
    print("DEBUG: User is in 'Working Hours Summary' tab.")

elif tab_option == "Violations":
    st.subheader("Violations")
    st.write("Placeholder for violations filtering.")
    print("DEBUG: User is in 'Violations' tab.")

else:
    st.subheader("Update Data")
    print("DEBUG: User is in 'Update Data' tab.")

    # Step 1: Filter by Main project and Month
    print("DEBUG: Building filter lists for 'Main project' and 'Month'.")
    projects_col = working_details["Main project"].dropna().astype(str).unique()
    all_projects = sorted(projects_col)
    months_col = working_details["Month"].dropna().astype(str).unique()
    all_months = sorted(months_col)
    with st.form("update_filter_form"):
        sel_projects = st.multiselect("Select Main Project(s)", options=all_projects, default=all_projects)
        sel_months = st.multiselect("Select Month(s)", options=all_months, default=all_months)
        filter_update = st.form_submit_button("Apply Filters")
    if filter_update:
        print("DEBUG: Filter form submitted.")
        df_update = working_details.copy()
        df_update["Main project"] = df_update["Main project"].astype(str)
        df_update["Month"] = df_update["Month"].astype(str)
        if sel_projects:
            print("DEBUG: Filtering by selected projects:", sel_projects)
            df_update = df_update[df_update["Main project"].isin(sel_projects)]
        if sel_months:
            print("DEBUG: Filtering by selected months:", sel_months)
            df_update = df_update[df_update["Month"].isin(sel_months)]

        st.dataframe(df_update, use_container_width=True)
        print("DEBUG: df_update shape after filtering:", df_update.shape)

        # Group by (Main project, Month)
        groups = df_update.groupby(["Main project", "Month"])
        print("DEBUG: Number of groups:", len(groups))

        # Step 2: Automatic or Manual
        mode = st.radio("Select Update Mode", ["Automatic", "Manual"], index=0)
        print("DEBUG: Update mode chosen:", mode)
        update_instructions = {}

        if mode == "Automatic":
            st.markdown("### Automatic Mode")
            print("DEBUG: Entering Automatic mode loop.")
            cat_fields = [
                "Functional Area (CRIT, CRIT - Data Management, CRIT - Data Governance, CRIT - Regulatory Reporting, CRIT - Portfolio Reporting, CRIT - Transformation)",
                "Project Category (Data Infrastructure, Monitoring & Insights, Analytics / Strategy Development, GDA Related, Trainings and Team Meeting)",
                "Complexity (H,M,L)",
                "Novelity (BAU repetitive, One time repetitive, New one time)",
                "Output Type (Core production work, Ad-hoc long-term projects, Ad-hoc short-term projects, Business Management, Administration, Trainings/L&D activities, Others) :",
                "Impact type (Customer Experience, Financial impact, Insights, Risk reduction, Others)"
            ]
            for (proj, mon), group in groups:
                with st.expander(f"Group: {proj} | {mon}", expanded=True):
                    print(f"DEBUG: Processing group {proj}, {mon} in Automatic mode.")
                    auto_start = pd.to_datetime(group["Start Date"], errors="coerce").min()
                    auto_start_str = auto_start.strftime("%m-%d-%Y") if pd.notna(auto_start) else ""
                    auto_comp = None
                    if "Completion Date" in group.columns:
                        auto_comp = pd.to_datetime(group["Completion Date"], errors="coerce").max()
                    auto_comp_str = auto_comp.strftime("%m-%d-%Y") if auto_comp is not None and pd.notna(auto_comp) else ""
                    cat_choices = {}
                    for cf in cat_fields:
                        if cf in group.columns and not group[cf].dropna().empty:
                            sorted_group = group.sort_values("RowNumber")
                            first_occ = sorted_group[cf].iloc[0]
                            most_freq = group[cf].mode().iloc[0]
                        else:
                            first_occ, most_freq = "", ""
                        choice = st.radio(
                            f"For {cf} in group {proj}/{mon}:",
                            ["First occurrence within month", "Most frequent within month"],
                            key=f"{proj}_{mon}_{cf}"
                        )
                        cat_val = first_occ if choice == "First occurrence within month" else most_freq
                        cat_choices[cf] = cat_val
                    update_instructions[f"{proj}||{mon}"] = {
                        "Start Date": auto_start_str,
                        "Completion Date": auto_comp_str,
                        **cat_choices
                    }
                    print(f"DEBUG: Built instructions for group {proj}/{mon} -> {update_instructions[f'{proj}||{mon}']}")

        else:
            st.markdown("### Manual Mode")
            print("DEBUG: Entering Manual mode loop.")
            allowed_values_manual = {
                # ... define your allowed options for each cat field ...
                "Functional Area (CRIT, CRIT - Data Management, CRIT - Data Governance, CRIT - Regulatory Reporting, CRIT - Portfolio Reporting, CRIT - Transformation)":
                    ["CRIT", "CRIT - Data Management", "CRIT - Data Governance", "CRIT - Regulatory Reporting", "CRIT - Portfolio Reporting", "CRIT - Transformation"],
                # etc...
            }
            for (proj, mon), group in groups:
                with st.expander(f"Group: {proj} | {mon}", expanded=True):
                    print(f"DEBUG: Processing group {proj}, {mon} in Manual mode.")
                    min_start = pd.to_datetime(group["Start Date"], errors="coerce").min()
                    max_comp = None
                    if "Completion Date" in group.columns:
                        max_comp = pd.to_datetime(group["Completion Date"], errors="coerce").max()
                    min_start_str = min_start.strftime("%m-%d-%Y") if pd.notna(min_start) else ""
                    max_comp_str = max_comp.strftime("%m-%d-%Y") if max_comp is not None and pd.notna(max_comp) else ""
                    new_start_str = st.text_input(f"Start Date for {proj}/{mon}", value=min_start_str, key=f"{proj}_{mon}_start")
                    new_comp_str = st.text_input(f"Completion Date for {proj}/{mon}", value=max_comp_str, key=f"{proj}_{mon}_comp")
                    manual_map = {
                        "Start Date": new_start_str,
                        "Completion Date": new_comp_str
                    }
                    # For each cat field
                    for cf, opts in allowed_values_manual.items():
                        if cf in group.columns and not group[cf].dropna().empty:
                            current_val = group[cf].iloc[0]
                        else:
                            current_val = ""
                        idx_default = opts.index(current_val) if current_val in opts else 0
                        chosen_val = st.selectbox(f"{cf} for {proj}/{mon}", options=opts, index=idx_default, key=f"{proj}_{mon}_{cf}")
                        manual_map[cf] = chosen_val
                    update_instructions[f"{proj}||{mon}"] = manual_map
                    print(f"DEBUG: Built instructions for group {proj}/{mon} -> {manual_map}")

        st.markdown("#### Final Update Instructions:")
        st.json(update_instructions)

        if st.button("Update Data"):
            print("DEBUG: 'Update Data' button clicked. Writing instructions to file, then updating Excel.")
            # Save instructions to text file
            try:
                with open(TEMP_UPDATE_FILE, "w", encoding="utf-8") as f:
                    json.dump(update_instructions, f, indent=2)
                st.success(f"Update instructions saved to {TEMP_UPDATE_FILE}.")
                print("DEBUG: Wrote instructions to file:", update_instructions)
            except Exception as e:
                print(f"DEBUG: Error writing to text file: {e}")
                st.error(f"Error writing to text file: {e}")
                st.stop()

            # Now update Excel from the text file
            try:
                wb = load_workbook(FILE_PATH)
                print("DEBUG: Opened workbook successfully.")
            except Exception as e:
                print(f"DEBUG: Error opening workbook: {e}")
                st.error(f"Error opening workbook: {e}")
                st.stop()

            updated_count = 0
            for i, row in working_details.iterrows():
                proj_val = str(row["Main project"])
                mon_val = str(row["Month"])
                key = f"{proj_val}||{mon_val}"
                if key in update_instructions:
                    instructions = update_instructions[key]
                    sheet_name = row["Employee"]
                    print(f"DEBUG: Checking row {i}, UniqueID={row['UniqueID']} => key={key} => sheet={sheet_name}")
                    if sheet_name not in wb.sheetnames:
                        print(f"DEBUG: Sheet {sheet_name} not found, skipping rowNumber={row['RowNumber']}.")
                        continue
                    ws = wb[sheet_name]
                    headers = {cell.value: cell.column for cell in ws[1]}
                    r_num = row["RowNumber"]
                    if not isinstance(r_num, int) or r_num < 1:
                        print(f"DEBUG: Invalid RowNumber {r_num} for UniqueID={row['UniqueID']}, skipping.")
                        continue

                    print(f"DEBUG: Updating row {r_num} in sheet {sheet_name} with instructions {instructions}")
                    # Start Date
                    if "Start Date" in headers and instructions.get("Start Date", ""):
                        ws.cell(row=r_num, column=headers["Start Date"], value=instructions["Start Date"])
                    # Completion Date
                    if "Completion Date" in headers and instructions.get("Completion Date", ""):
                        ws.cell(row=r_num, column=headers["Completion Date"], value=instructions["Completion Date"])
                    # For each categorical col
                    # (Adapt these if you have the full dictionary)
                    for cf in [
                        "Functional Area (CRIT, CRIT - Data Management, CRIT - Data Governance, CRIT - Regulatory Reporting, CRIT - Portfolio Reporting, CRIT - Transformation)",
                        "Project Category (Data Infrastructure, Monitoring & Insights, Analytics / Strategy Development, GDA Related, Trainings and Team Meeting)",
                        "Complexity (H,M,L)",
                        "Novelity (BAU repetitive, One time repetitive, New one time)",
                        "Output Type (Core production work, Ad-hoc long-term projects, Ad-hoc short-term projects, Business Management, Administration, Trainings/L&D activities, Others) :",
                        "Impact type (Customer Experience, Financial impact, Insights, Risk reduction, Others)"
                    ]:
                        if cf in headers and instructions.get(cf, ""):
                            ws.cell(row=r_num, column=headers[cf], value=instructions[cf])
                    updated_count += 1
            print(f"DEBUG: Total rows updated: {updated_count}")

            try:
                wb.save(FILE_PATH)
                st.success("Excel file updated successfully.")
                print("DEBUG: Workbook saved successfully. Updated rows:", updated_count)
            except Exception as e:
                print(f"DEBUG: Error saving workbook: {e}")
                st.error(f"Error saving workbook: {e}")
    else:
        print("DEBUG: No filter form submission yet, so no update logic triggered.")
        st.info("Apply filters above to see update data.")