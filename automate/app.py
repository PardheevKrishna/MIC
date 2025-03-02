import streamlit as st
import pandas as pd
from datetime import timedelta
from io import BytesIO
from openpyxl import load_workbook
import json
import os

# ---------- CONFIGURATION ----------
FILE_PATH = "input.xlsx"  # CHANGE this to your actual Excel file path

# -------------- CUSTOM CSS & TITLE --------------
st.markdown("""
<style>
.mySkipButton { visibility: hidden; }
</style>
""", unsafe_allow_html=True)
st.title("Team Report Dashboard")

# ===================== PROCESS FUNCTION =====================
def process_excel_file(file_path):
    """
    Reads each employee sheet (employee names from "Home") and returns two DataFrames:
      - working_details: row-level data for all employees
      - violations_df: flagged violations (Invalid value, Start date change, Weekly < 40, etc.)
    """
    allowed_values = {
        "Functional Area (CRIT, CRIT - Data Management, CRIT - Data Governance, CRIT - Regulatory Reporting, CRIT - Portfolio Reporting, CRIT - Transformation)": [
            "CRIT", "CRIT - Data Management", "CRIT - Data Governance", "CRIT - Regulatory Reporting", "CRIT - Portfolio Reporting", "CRIT - Transformation"
        ],
        "Project Category (Data Infrastructure, Monitoring & Insights, Analytics / Strategy Development, GDA Related, Trainings and Team Meeting)": [
            "Data Infrastructure", "Monitoring & Insights", "Analytics / Strategy Development", "GDA Related", "Trainings and Team Meeting"
        ],
        "Complexity (H,M,L)": ["H", "M", "L"],
        "Novelity (BAU repetitive, One time repetitive, New one time)": [
            "BAU repetitive", "One time repetitive", "New one time"
        ],
        "Output Type (Core production work, Ad-hoc long-term projects, Ad-hoc short-term projects, Business Management, Administration, Trainings/L&D activities, Others) :": [
            "Core production work", "Ad-hoc long-term projects", "Ad-hoc short-term projects", "Business Management", "Administration", "Trainings/L&D activities", "Others"
        ],
        "Impact type (Customer Experience, Financial impact, Insights, Risk reduction, Others)": [
            "Customer Experience", "Financial impact", "Insights", "Risk reduction", "Others"
        ]
    }
    start_date_exceptions = [
        "Internal meetings", "Internal Meetings", "Internal meeting", "internal meeting",
        "External meetings", "External Meeting", "External meeting", "external meetings",
        "Sick leave", "Sick Leave", "Sick day",
        "Annual meeting", "annual meeting", "Traveling", "Develop/Dev training",
        "Internal Taining", "internal taining", "Interview"
    ]
    try:
        home_df = pd.read_excel(file_path, sheet_name="Home", header=None)
    except Exception as e:
        st.error(f"Error reading Home sheet: {e}")
        return None, None

    employee_names = home_df.iloc[2:, 5].dropna().astype(str).tolist()
    xls = pd.ExcelFile(file_path)
    all_sheet_names = xls.sheet_names

    working_list = []
    viol_list = []  # list to collect violation dictionaries
    project_month_info = {}

    for emp in employee_names:
        if emp not in all_sheet_names:
            continue
        try:
            df = pd.read_excel(file_path, sheet_name=emp)
        except Exception as e:
            st.warning(f"Could not read sheet for {emp}: {e}")
            continue

        df.columns = [str(c).replace("\n", " ").strip() for c in df.columns]
        required_cols = ["Status Date (Every Friday)", "Main project", "Name of the Project", "Start Date", "Weekly Time Spent(Hrs)"]
        if not all(c in df.columns for c in required_cols):
            continue

        df["Employee"] = emp
        df["RowNumber"] = df.index + 2  # Excel row (header is row 1)
        df["Status Date (Every Friday)"] = pd.to_datetime(
            df["Status Date (Every Friday)"], format="%m-%d-%Y", errors="coerce"
        )

        # (1) Validate allowed values
        for col, a_list in allowed_values.items():
            if col not in df.columns:
                continue
            for i, val in df[col].items():
                if pd.isna(val):
                    continue
                tokens = [t.strip() for t in str(val).split(",") if t.strip()]
                if len(tokens) != 1 or tokens[0] not in a_list:
                    viol_list.append({
                        "Employee": emp,
                        "Violation Type": "Invalid value",
                        "Violation Details": f"{col} = {val}",
                        "Location": f"Sheet {emp}, Row {df.at[i, 'RowNumber']}",
                        "Violation Date": df.at[i, "Status Date (Every Friday)"]
                    })

        # (2) Check start date consistency (skip exceptions)
        for i, row in df.iterrows():
            proj = row["Name of the Project"]
            start_val = row["Start Date"]
            mp_val = str(row["Main project"]).strip() if pd.notna(row["Main project"]) else ""
            proj_val = str(proj).strip() if pd.notna(proj) else ""
            if mp_val in start_date_exceptions or proj_val in start_date_exceptions:
                continue
            if pd.notna(proj) and pd.notna(start_val) and pd.notna(row["Status Date (Every Friday)"]):
                month_key = row["Status Date (Every Friday)"].strftime("%Y-%m")
                key = (proj, month_key)
                current_start = pd.to_datetime(start_val, format="%m-%d-%Y", errors="coerce")
                if key not in project_month_info:
                    project_month_info[key] = current_start
                else:
                    baseline = project_month_info[key]
                    if pd.notna(current_start) and pd.notna(baseline) and current_start != baseline:
                        old_str = baseline.strftime("%m-%d-%Y")
                        new_str = current_start.strftime("%m-%d-%Y")
                        viol_list.append({
                            "Employee": emp,
                            "Violation Type": "Start date change",
                            "Violation Details": f"{proj}: expected {old_str}, got {new_str}",
                            "Location": f"Sheet {emp}, Row {row['RowNumber']}",
                            "Violation Date": row["Status Date (Every Friday)"]
                        })

        # (3) Weekly hours check
        df["Weekly Time Spent(Hrs)"] = pd.to_numeric(df["Weekly Time Spent(Hrs)"], errors="coerce").fillna(0)
        friday_dates = df[(df["Status Date (Every Friday)"].dt.weekday == 4) & (df["Status Date (Every Friday)"].notna())]["Status Date (Every Friday)"].unique()
        for friday in friday_dates:
            if pd.isna(friday):
                continue
            friday_str = friday.strftime("%m-%d-%Y")
            week_start = friday - timedelta(days=4)
            week_df = df[(df["Status Date (Every Friday)"] >= week_start) & (df["Status Date (Every Friday)"] <= friday)]
            total_hrs = week_df["Weekly Time Spent(Hrs)"].sum()
            if total_hrs < 40:
                row_nums_str = ", ".join(str(x) for x in week_df["RowNumber"].tolist())
                viol_list.append({
                    "Employee": emp,
                    "Violation Type": "Working hours less than 40",
                    "Violation Details": f"Week ending {friday_str} insufficient hours",
                    "Location": f"Sheet {emp}, Rows: {row_nums_str}",
                    "Violation Date": friday
                })

        # (4) Additional columns for later use
        df["PTO Hours"] = df.apply(lambda r: r["Weekly Time Spent(Hrs)"] if "PTO" in str(r["Main project"]) else 0, axis=1)
        df["Work Hours"] = df.apply(lambda r: r["Weekly Time Spent(Hrs)"] if "PTO" not in str(r["Main project"]) else 0, axis=1)
        df["Month"] = df["Status Date (Every Friday)"].dt.to_period("M").astype(str)
        df["WeekFriday"] = df["Status Date (Every Friday)"].dt.strftime("%m-%d-%Y").fillna("N/A")
        df["UniqueID"] = df["Employee"] + "_" + df["RowNumber"].astype(str)

        working_list.append(df)

    working_details = pd.concat(working_list, ignore_index=True) if working_list else pd.DataFrame()
    violations_df = pd.DataFrame(viol_list)
    return working_details, violations_df

# ======= SESSION STATE DATA LOADING / INITIALIZATION =======
if "working_details" not in st.session_state or "violations_df" not in st.session_state:
    wd, vd = process_excel_file(FILE_PATH)
    if wd is None or vd is None:
        st.error("Error processing the Excel file.")
        st.stop()
    st.session_state["working_details"] = wd
    st.session_state["violations_df"] = vd
    st.success("Reports generated successfully!")
else:
    st.success("Using cached data from session state.")

working_details = st.session_state["working_details"]
violations_df = st.session_state["violations_df"]

# ---------- USE st.radio TO SELECT A TAB (this preserves the active view) ----------
tab_option = st.radio("Select Report Tab", 
                       ["Team Monthly Summary", "Working Hours Summary", "Violations and Update"],
                       index=0)

# ---------- TEAM MONTHLY SUMMARY (Tab 1) ----------
if tab_option == "Team Monthly Summary":
    st.subheader("Team Monthly Summary")
    # [Insert your Team Monthly Summary code here]
    st.write("Team Monthly Summary code goes here.")

# ---------- WORKING HOURS SUMMARY (Tab 2) ----------
elif tab_option == "Working Hours Summary":
    st.subheader("Working Hours Summary")
    # [Insert your Working Hours Summary code here]
    st.write("Working Hours Summary code goes here.")

# ---------- VIOLATIONS & UPDATE (Tab 3) ----------
elif tab_option == "Violations and Update":
    st.subheader("Violations and Update")
    if violations_df.empty:
        st.info("No violations found.")
    else:
        # Step 1: Filter Violations
        all_emps_v = sorted(violations_df["Employee"].dropna().unique())
        all_types_v = ["Invalid value", "Working hours less than 40", "Start date change"]

        with st.form("violations_filter_form"):
            col1_v, col2_v = st.columns([0.7, 0.3])
            emp_sel_v = col1_v.multiselect("Select Employee(s)", options=all_emps_v)
            select_all_emp_v = col2_v.checkbox("Select All Employees")
            col3_v, col4_v = st.columns([0.7, 0.3])
            type_sel_v = col3_v.multiselect("Select Violation Type(s)", options=all_types_v)
            select_all_type_v = col4_v.checkbox("Select All Violation Types")
            filter_btn_v = st.form_submit_button("Filter Violations")
        if filter_btn_v:
            if select_all_emp_v:
                emp_sel_v = all_emps_v
            if select_all_type_v:
                type_sel_v = all_types_v
            df_v = violations_df.copy()
            if emp_sel_v:
                df_v = df_v[df_v["Employee"].isin(emp_sel_v)]
            if type_sel_v:
                df_v = df_v[df_v["Violation Type"].isin(type_sel_v)]
            st.dataframe(df_v, use_container_width=True)
            
            # Step 2: Row Selection (this part no longer forces a tab change)
            all_ids = sorted(df_v.apply(lambda r: f"{r['Employee']}_{r['Location'].split('Row ')[-1]}", axis=1).unique())
            st.markdown("#### Select Rows to Update (by UniqueID)")
            select_all_rows = st.checkbox("Select All Rows", key="select_all_rows")
            if select_all_rows:
                selected_ids = all_ids
            else:
                selected_ids = st.multiselect("Select UniqueIDs", options=all_ids, key="selected_ids")
            
            # Button to load editing forms – this does not trigger a page change because we use radio for tabs
            if st.button("Load Editing Forms", key="load_editing"):
                if not selected_ids:
                    st.error("No rows selected for update.")
                else:
                    st.session_state["selected_rows"] = selected_ids
                    st.success(f"Selected rows: {selected_ids}")

        # Step 3: Load Editing Forms (only if row selection has been confirmed)
        if "selected_rows" in st.session_state:
            st.markdown("### Edit Each Selected Row")
            updated_data = {}  # Dictionary to collect new values
            cat_fields = [
                "Functional Area (CRIT, CRIT - Data Management, CRIT - Data Governance, CRIT - Regulatory Reporting, CRIT - Portfolio Reporting, CRIT - Transformation)",
                "Project Category (Data Infrastructure, Monitoring & Insights, Analytics / Strategy Development, GDA Related, Trainings and Team Meeting)",
                "Complexity (H,M,L)",
                "Novelity (BAU repetitive, One time repetitive, New one time)",
                "Output Type (Core production work, Ad-hoc long-term projects, Ad-hoc short-term projects, Business Management, Administration, Trainings/L&D activities, Others) :",
                "Impact type (Customer Experience, Financial impact, Insights, Risk reduction, Others)"
            ]
            # Build a lookup from UniqueID to corresponding row in working_details
            working_details_dict = {}
            for idx, row in working_details.iterrows():
                uid = row["UniqueID"]
                working_details_dict[uid] = row
            
            for uid in st.session_state["selected_rows"]:
                if uid not in working_details_dict:
                    st.warning(f"No data found for {uid}")
                    continue
                row = working_details_dict[uid]
                with st.expander(f"Edit row {uid}", expanded=True):
                    new_start = st.text_input("Start Date", value=str(row["Start Date"]), key=f"{uid}_start")
                    new_comp = st.text_input("Completion Date", value=str(row.get("Completion Date", "")), key=f"{uid}_comp")
                    row_data = {
                        "Employee": row["Employee"],
                        "RowNumber": row["RowNumber"],
                        "Start Date": new_start,
                        "Completion Date": new_comp
                    }
                    for cf in cat_fields:
                        row_data[cf] = st.text_input(cf, value=str(row.get(cf, "")), key=f"{uid}_{cf}")
                    updated_data[uid] = row_data

            # Step 4: Update Excel using a pure Python function (no Streamlit involvement in the update logic)
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
                        headers = {cell.value: cell.column for cell in ws[1]}
                        r_num = row_vals["RowNumber"]
                        if "Start Date" in headers and row_vals["Start Date"]:
                            ws.cell(row=r_num, column=headers["Start Date"], value=row_vals["Start Date"])
                        if "Completion Date" in headers and row_vals["Completion Date"]:
                            ws.cell(row=r_num, column=headers["Completion Date"], value=row_vals["Completion Date"])
                        for cf in cat_fields:
                            if cf in headers and row_vals.get(cf, ""):
                                ws.cell(row=r_num, column=headers[cf], value=row_vals[cf])
                    try:
                        wb.save(FILE_PATH)
                        return True
                    except Exception as e:
                        st.error(f"Error saving workbook: {e}")
                        return False

                if update_excel_file_py(updated_data):
                    st.success("Excel file updated successfully (via pure Python function).")
                    # Optionally update session_state working_details so that changes show immediately
                    for uid, row_vals in updated_data.items():
                        mask = st.session_state["working_details"]["UniqueID"] == uid
                        if mask.any():
                            idx = st.session_state["working_details"].index[mask][0]
                            for col, new_val in row_vals.items():
                                st.session_state["working_details"].at[idx, col] = new_val
                    st.success("Session data updated. Changes now appear in your dashboard.")