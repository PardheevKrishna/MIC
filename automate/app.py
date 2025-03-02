import streamlit as st
import pandas as pd
import json
from datetime import datetime, timedelta
from io import BytesIO
from openpyxl import load_workbook
import os

# ---------- CONFIGURATION ----------
FILE_PATH = "input.xlsx"        # Change this to your actual Excel file path
TEMP_JSON_FILE = "temp_changes.json"  # Where we temporarily store row edits

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
      - violations_df: flagged violations
    """
    # Allowed categorical values
    allowed_values = {
        "Functional Area (CRIT, CRIT - Data Management, CRIT - Data Governance, CRIT - Regulatory Reporting, CRIT - Portfolio Reporting, CRIT - Transformation)": [
            "CRIT", "CRIT - Data Management", "CRIT - Data Governance", "CRIT - Regulatory Reporting", "CRIT - Portfolio Reporting", "CRIT - Transformation"
        ],
        "Project Category (Data Infrastructure, Monitoring & Insights, Analytics / Strategy Development, GDA Related, Trainings and Team Meeting)": [
            "Data Infrastructure", "Monitoring & Insights", "Analytics / Strategy Development", "GDA Related", "Trainings and Team Meeting"
        ],
        "Complexity (H,M,L)": [
            "H", "M", "L"
        ],
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

    # Start date exceptions
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
    viol_list = []
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
        req_cols = ["Status Date (Every Friday)", "Main project", "Name of the Project", "Start Date", "Weekly Time Spent(Hrs)"]
        if not all(c in df.columns for c in req_cols):
            continue

        df["Employee"] = emp
        df["RowNumber"] = df.index + 2
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

        # (2) Start Date consistency
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

        # (4) Additional columns
        df["PTO Hours"] = df.apply(lambda r: r["Weekly Time Spent(Hrs)"] if "PTO" in str(r["Main project"]) else 0, axis=1)
        df["Work Hours"] = df.apply(lambda r: r["Weekly Time Spent(Hrs)"] if "PTO" not in str(r["Main project"]) else 0, axis=1)
        df["Month"] = df["Status Date (Every Friday)"].dt.to_period("M").astype(str)
        df["WeekFriday"] = df["Status Date (Every Friday)"].dt.strftime("%m-%d-%Y").fillna("N/A")
        # Unique ID
        df["UniqueID"] = df["Employee"] + "_" + df["RowNumber"].astype(str)

        working_list.append(df)

    if working_list:
        working_details = pd.concat(working_list, ignore_index=True)
    else:
        working_details = pd.DataFrame()

    violations_df = pd.DataFrame(viol_list)
    return working_details, violations_df

# ---- LOAD DATA ----
wd, vd = process_excel_file(FILE_PATH)
if wd is None or vd is None:
    st.error("Error processing the Excel file.")
    st.stop()
else:
    st.success("Reports generated successfully!")

# TABS
tab_option = st.radio("Select a Tab", ["Team Monthly Summary", "Working Hours Summary", "Violations and Update"])

if tab_option == "Team Monthly Summary":
    st.subheader("Team Monthly Summary")
    st.write("Placeholder for your summary logic.")

elif tab_option == "Working Hours Summary":
    st.subheader("Working Hours Summary")
    st.write("Placeholder for your summary logic.")

else:
    st.subheader("Violations and Update (No Session State)")
    if vd.empty:
        st.info("No violations found.")
    else:
        # Step A: Filter
        all_emps_v = sorted(vd["Employee"].dropna().unique())
        all_types_v = ["Invalid value", "Working hours less than 40", "Start date change"]
        with st.form("violations_filter_form"):
            col1, col2 = st.columns([0.7, 0.3])
            emp_sel_v = col1.multiselect("Select Employee(s)", options=all_emps_v)
            sel_all_emp = col2.checkbox("Select All Employees")
            col3, col4 = st.columns([0.7, 0.3])
            type_sel_v = col3.multiselect("Select Violation Type(s)", options=all_types_v)
            sel_all_type = col4.checkbox("Select All Types")
            filter_btn = st.form_submit_button("Filter Violations")

        if filter_btn:
            if sel_all_emp:
                emp_sel_v = all_emps_v
            if sel_all_type:
                type_sel_v = all_types_v
            df_v = vd.copy()
            if emp_sel_v:
                df_v = df_v[df_v["Employee"].isin(emp_sel_v)]
            if type_sel_v:
                df_v = df_v[df_v["Violation Type"].isin(type_sel_v)]
            st.dataframe(df_v, use_container_width=True)

            # Step B: Row Selection
            all_ids = sorted(df_v["UniqueID"].unique())
            st.markdown("#### Select Rows to Update")
            select_all_rows = st.checkbox("Select All Rows")
            if select_all_rows:
                selected_ids = all_ids
            else:
                selected_ids = st.multiselect("UniqueIDs", options=all_ids)

            # Step C: Choose Automatic or Manual
            update_mode = st.radio("Update Mode", ["Automatic", "Manual"], index=0)
            load_form_btn = st.button("Load Editing Form")

            if load_form_btn:
                if not selected_ids:
                    st.error("No rows selected.")
                else:
                    st.write(f"Rows selected: {selected_ids}")
                    st.markdown("### Edit Each Row")

                    # Build a map from UniqueID -> row in wd
                    wd_map = {}
                    for i, r in wd.iterrows():
                        wd_map[r["UniqueID"]] = r

                    updated_rows = {}
                    cat_fields = [
                        "Functional Area (CRIT, CRIT - Data Management, CRIT - Data Governance, CRIT - Regulatory Reporting, CRIT - Portfolio Reporting, CRIT - Transformation)",
                        "Project Category (Data Infrastructure, Monitoring & Insights, Analytics / Strategy Development, GDA Related, Trainings and Team Meeting)",
                        "Complexity (H,M,L)",
                        "Novelity (BAU repetitive, One time repetitive, New one time)",
                        "Output Type (Core production work, Ad-hoc long-term projects, Ad-hoc short-term projects, Business Management, Administration, Trainings/L&D activities, Others) :",
                        "Impact type (Customer Experience, Financial impact, Insights, Risk reduction, Others)"
                    ]

                    def compute_auto_suggestions(row):
                        # group by same "Main project"
                        mp = row["Main project"]
                        group = wd[wd["Main project"] == mp]
                        auto_start = pd.to_datetime(group["Start Date"], errors="coerce").min()
                        auto_comp = None
                        if "Completion Date" in group.columns:
                            auto_comp = pd.to_datetime(group["Completion Date"], errors="coerce").max()
                        if pd.notna(auto_start):
                            auto_start_str = auto_start.strftime("%m-%d-%Y")
                        else:
                            auto_start_str = ""
                        if auto_comp is not None and pd.notna(auto_comp):
                            auto_comp_str = auto_comp.strftime("%m-%d-%Y")
                        else:
                            auto_comp_str = ""
                        cat_sugg = {}
                        for cf in cat_fields:
                            if cf in group.columns and not group[cf].dropna().empty:
                                cat_sugg[cf] = group[cf].mode().iloc[0]
                            else:
                                cat_sugg[cf] = ""
                        return auto_start_str, auto_comp_str, cat_sugg

                    for uid in selected_ids:
                        row = wd_map.get(uid, None)
                        if row is None:
                            st.warning(f"No row found for {uid} in working_details")
                            continue
                        with st.expander(f"Edit row: {uid}", expanded=True):
                            if update_mode == "Automatic":
                                s_str, c_str, cat_map = compute_auto_suggestions(row)
                                new_start = st.text_input("Start Date", value=s_str, key=f"{uid}_start")
                                new_comp = st.text_input("Completion Date", value=c_str, key=f"{uid}_comp")
                                new_cats = {}
                                for cf in cat_fields:
                                    new_cats[cf] = st.text_input(cf, value=cat_map[cf], key=f"{uid}_{cf}")
                                updated_rows[uid] = {
                                    "Employee": row["Employee"],
                                    "RowNumber": row["RowNumber"],
                                    "Start Date": new_start,
                                    "Completion Date": new_comp,
                                    **new_cats
                                }
                            else:
                                # Manual
                                start_val = str(row.get("Start Date",""))
                                comp_val = str(row.get("Completion Date",""))
                                new_start = st.text_input("Start Date", value=start_val, key=f"{uid}_start")
                                new_comp = st.text_input("Completion Date", value=comp_val, key=f"{uid}_comp")
                                new_cats = {}
                                for cf in cat_fields:
                                    cur_val = str(row.get(cf,""))
                                    new_cats[cf] = st.text_input(cf, value=cur_val, key=f"{uid}_{cf}")
                                updated_rows[uid] = {
                                    "Employee": row["Employee"],
                                    "RowNumber": row["RowNumber"],
                                    "Start Date": new_start,
                                    "Completion Date": new_comp,
                                    **new_cats
                                }

                    # Step D: Save to text file
                    if st.button("Save to Text File"):
                        try:
                            with open(TEMP_JSON_FILE, "w", encoding="utf-8") as f:
                                json.dump(updated_rows, f, indent=2)
                            st.success(f"Edits saved to {TEMP_JSON_FILE}.")
                        except Exception as e:
                            st.error(f"Error writing to text file: {e}")

                    # Step E: Update Excel from text file
                    if st.button("Update Excel from Text File"):
                        if not os.path.exists(TEMP_JSON_FILE):
                            st.error(f"No file {TEMP_JSON_FILE} found. Please save your changes first.")
                        else:
                            try:
                                with open(TEMP_JSON_FILE, "r", encoding="utf-8") as f:
                                    changes = json.load(f)
                            except Exception as e:
                                st.error(f"Error reading from text file: {e}")
                                st.stop()

                            # Apply changes
                            try:
                                wb = load_workbook(FILE_PATH)
                            except Exception as e:
                                st.error(f"Error opening workbook: {e}")
                                st.stop()

                            for uid, row_vals in changes.items():
                                sheet_name = row_vals["Employee"]
                                if sheet_name not in wb.sheetnames:
                                    st.warning(f"Sheet {sheet_name} not found in Excel.")
                                    continue
                                ws = wb[sheet_name]
                                # Build header map
                                headers = {cell.value: cell.column for cell in ws[1]}
                                r_num = row_vals["RowNumber"]
                                # If columns exist, update them
                                if "Start Date" in headers and row_vals["Start Date"]:
                                    ws.cell(row=r_num, column=headers["Start Date"], value=row_vals["Start Date"])
                                if "Completion Date" in headers and row_vals["Completion Date"]:
                                    ws.cell(row=r_num, column=headers["Completion Date"], value=row_vals["Completion Date"])
                                for cf in cat_fields:
                                    if cf in headers and row_vals.get(cf,""):
                                        ws.cell(row=r_num, column=headers[cf], value=row_vals[cf])

                            try:
                                wb.save(FILE_PATH)
                                st.success("Excel file updated successfully (from text file).")
                            except Exception as e:
                                st.error(f"Error saving workbook: {e}")
        else:
            st.info("Use the form above to filter violations.")