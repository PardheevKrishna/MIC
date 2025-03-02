import streamlit as st
import pandas as pd
import json
from datetime import datetime, timedelta
from openpyxl import load_workbook
from io import BytesIO
import os

# ---------- CONFIGURATION ----------
FILE_PATH = "input.xlsx"  # Change this to your actual Excel file path
TEMP_UPDATE_FILE = "temp_update.json"  # Temporary file to store update instructions

st.set_page_config(page_title="Team Report Dashboard", layout="wide")
st.title("Team Report Dashboard")

# ===================== PROCESS EXCEL FILE =====================
def process_excel_file(file_path):
    """
    Reads each employee sheet (employee names in "Home") and returns two DataFrames:
      - working_details: All rows from each employee sheet with added columns:
            Employee, RowNumber, Month (yyyy-mm), WeekFriday (mm-dd-yyyy)
      - violations_df: Rows flagged as violations (not used in Update Data but kept for context)
    """
    # Allowed categorical values with proper spacing
    allowed_values = {
        "Functional Area (CRIT, CRIT - Data Management, CRIT - Data Governance, CRIT - Regulatory Reporting, CRIT - Portfolio Reporting, CRIT - Transformation)":
            ["CRIT", "CRIT - Data Management", "CRIT - Data Governance", "CRIT - Regulatory Reporting", "CRIT - Portfolio Reporting", "CRIT - Transformation"],
        "Project Category (Data Infrastructure, Monitoring & Insights, Analytics / Strategy Development, GDA Related, Trainings and Team Meeting)":
            ["Data Infrastructure", "Monitoring & Insights", "Analytics / Strategy Development", "GDA Related", "Trainings and Team Meeting"],
        "Complexity (H,M,L)":
            ["H", "M", "L"],
        "Novelity (BAU repetitive, One time repetitive, New one time)":
            ["BAU repetitive", "One time repetitive", "New one time"],
        "Output Type (Core production work, Ad-hoc long-term projects, Ad-hoc short-term projects, Business Management, Administration, Trainings/L&D activities, Others) :":
            ["Core production work", "Ad-hoc long-term projects", "Ad-hoc short-term projects", "Business Management", "Administration", "Trainings/L&D activities", "Others"],
        "Impact type (Customer Experience, Financial impact, Insights, Risk reduction, Others)":
            ["Customer Experience", "Financial impact", "Insights", "Risk reduction", "Others"]
    }
    # Exceptions for Start Date check
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
    try:
        xls = pd.ExcelFile(file_path)
        all_sheet_names = xls.sheet_names
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        return None, None

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
        df["RowNumber"] = df.index + 2  # Header is row 1
        df["Status Date (Every Friday)"] = pd.to_datetime(df["Status Date (Every Friday)"], format="%m-%d-%Y", errors="coerce")

        # Validate allowed values (not critical for update, but included)
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

        # Check start date consistency within each project/month
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

        # Weekly hours check (omitted for brevity)
        df["Weekly Time Spent(Hrs)"] = pd.to_numeric(df["Weekly Time Spent(Hrs)"], errors="coerce").fillna(0)
        friday_dates = df[(df["Status Date (Every Friday)"].dt.weekday == 4) & (df["Status Date (Every Friday)"].notna())]["Status Date (Every Friday)"].unique()
        for friday in friday_dates:
            if pd.isna(friday):
                continue
            friday_str = friday.strftime("%m-%d-%Y")
            week_start = friday - timedelta(days=4)
            week_df = df[(df["Status Date (Every Friday)"] >= week_start) & (df["Status Date (Every Friday)"] <= friday)]
            if week_df["Weekly Time Spent(Hrs)"].sum() < 40:
                row_nums_str = ", ".join(str(x) for x in week_df["RowNumber"].tolist())
                viol_list.append({
                    "Employee": emp,
                    "Violation Type": "Working hours less than 40",
                    "Violation Details": f"Week ending {friday_str} insufficient hours",
                    "Location": f"Sheet {emp}, Rows: {row_nums_str}",
                    "Violation Date": friday
                })

        # Additional columns
        df["PTO Hours"] = df.apply(lambda r: r["Weekly Time Spent(Hrs)"] if "PTO" in str(r["Main project"]) else 0, axis=1)
        df["Work Hours"] = df.apply(lambda r: r["Weekly Time Spent(Hrs)"] if "PTO" not in str(r["Main project"]) else 0, axis=1)
        df["Month"] = df["Status Date (Every Friday)"].dt.to_period("M").astype(str)
        df["WeekFriday"] = df["Status Date (Every Friday)"].dt.strftime("%m-%d-%Y").fillna("N/A")
        df["UniqueID"] = df["Employee"] + "_" + df["RowNumber"].astype(str)

        working_list.append(df)

    working_details = pd.concat(working_list, ignore_index=True) if working_list else pd.DataFrame()
    violations_df = pd.DataFrame(viol_list)
    return working_details, violations_df

# ---- LOAD DATA ----
working_details, violations_df = process_excel_file(FILE_PATH)
if working_details is None:
    st.error("Error processing the Excel file.")
    st.stop()
else:
    st.success("Data loaded successfully!")

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

elif tab_option == "Working Hours Summary":
    st.subheader("Working Hours Summary")
    st.write("Placeholder for working hours logic.")

elif tab_option == "Violations":
    st.subheader("Violations")
    st.write("Placeholder for violations filtering.")

else:
    st.subheader("Update Data")

    # Step 1: Filter by Main project and Month
    # Convert to string to avoid TypeError when sorting
    projects_col = working_details["Main project"].dropna().astype(str).unique()
    all_projects = sorted(projects_col)
    months_col = working_details["Month"].dropna().astype(str).unique()
    all_months = sorted(months_col)

    with st.form("update_filter_form"):
        sel_projects = st.multiselect("Select Main Project(s)", options=all_projects, default=all_projects)
        sel_months = st.multiselect("Select Month(s)", options=all_months, default=all_months)
        filter_update = st.form_submit_button("Apply Filters")

    if filter_update:
        df_update = working_details.copy()
        # Convert df_update columns to string if needed
        df_update["Main project"] = df_update["Main project"].astype(str)
        df_update["Month"] = df_update["Month"].astype(str)

        if sel_projects:
            df_update = df_update[df_update["Main project"].isin(sel_projects)]
        if sel_months:
            df_update = df_update[df_update["Month"].isin(sel_months)]

        st.dataframe(df_update, use_container_width=True)

        # Group by (Main project, Month)
        groups = df_update.groupby(["Main project", "Month"])

        # Step 2: Automatic or Manual
        mode = st.radio("Select Update Mode", ["Automatic", "Manual"], index=0)
        update_instructions = {}

        if mode == "Automatic":
            st.markdown("### Automatic Mode")
            for (proj, mon), group in groups:
                with st.expander(f"Group: {proj} | {mon}", expanded=True):
                    # For Start Date: earliest in that group
                    auto_start = pd.to_datetime(group["Start Date"], errors="coerce").min()
                    auto_start_str = auto_start.strftime("%m-%d-%Y") if pd.notna(auto_start) else ""
                    # For Completion Date: latest in that group
                    auto_comp = None
                    if "Completion Date" in group.columns:
                        auto_comp = pd.to_datetime(group["Completion Date"], errors="coerce").max()
                    auto_comp_str = auto_comp.strftime("%m-%d-%Y") if auto_comp is not None and pd.notna(auto_comp) else ""
                    st.write(f"**Earliest Start:** {auto_start_str}")
                    st.write(f"**Latest Completion:** {auto_comp_str}")

                    # For each categorical col, user picks "first occurrence" or "most frequent"
                    cat_fields = [
                        "Functional Area (CRIT, CRIT - Data Management, CRIT - Data Governance, CRIT - Regulatory Reporting, CRIT - Portfolio Reporting, CRIT - Transformation)",
                        "Project Category (Data Infrastructure, Monitoring & Insights, Analytics / Strategy Development, GDA Related, Trainings and Team Meeting)",
                        "Complexity (H,M,L)",
                        "Novelity (BAU repetitive, One time repetitive, New one time)",
                        "Output Type (Core production work, Ad-hoc long-term projects, Ad-hoc short-term projects, Business Management, Administration, Trainings/L&D activities, Others) :",
                        "Impact type (Customer Experience, Financial impact, Insights, Risk reduction, Others)"
                    ]
                    cat_choices = {}
                    for cf in cat_fields:
                        if cf in group.columns and not group[cf].dropna().empty:
                            # first occurrence = the value from the earliest row
                            sorted_group = group.sort_values("RowNumber")
                            first_occ = sorted_group[cf].iloc[0]
                            most_freq = group[cf].mode().iloc[0]
                        else:
                            first_occ, most_freq = "", ""
                        choice = st.radio(
                            f"{cf}",
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

        else:
            st.markdown("### Manual Mode")
            allowed_values_manual = {
                "Functional Area (CRIT, CRIT - Data Management, CRIT - Data Governance, CRIT - Regulatory Reporting, CRIT - Portfolio Reporting, CRIT - Transformation)":
                    ["CRIT", "CRIT - Data Management", "CRIT - Data Governance", "CRIT - Regulatory Reporting", "CRIT - Portfolio Reporting", "CRIT - Transformation"],
                "Project Category (Data Infrastructure, Monitoring & Insights, Analytics / Strategy Development, GDA Related, Trainings and Team Meeting)":
                    ["Data Infrastructure", "Monitoring & Insights", "Analytics / Strategy Development", "GDA Related", "Trainings and Team Meeting"],
                "Complexity (H,M,L)": ["H", "M", "L"],
                "Novelity (BAU repetitive, One time repetitive, New one time)":
                    ["BAU repetitive", "One time repetitive", "New one time"],
                "Output Type (Core production work, Ad-hoc long-term projects, Ad-hoc short-term projects, Business Management, Administration, Trainings/L&D activities, Others) :":
                    ["Core production work", "Ad-hoc long-term projects", "Ad-hoc short-term projects", "Business Management", "Administration", "Trainings/L&D activities", "Others"],
                "Impact type (Customer Experience, Financial impact, Insights, Risk reduction, Others)":
                    ["Customer Experience", "Financial impact", "Insights", "Risk reduction", "Others"]
            }
            for (proj, mon), group in groups:
                with st.expander(f"Group: {proj} | {mon}", expanded=True):
                    # min Start, max Completion
                    min_start = pd.to_datetime(group["Start Date"], errors="coerce").min()
                    max_comp = pd.to_datetime(group["Completion Date"], errors="coerce").max() if "Completion Date" in group.columns else None
                    min_start_str = min_start.strftime("%m-%d-%Y") if pd.notna(min_start) else ""
                    max_comp_str = max_comp.strftime("%m-%d-%Y") if max_comp is not None and pd.notna(max_comp) else ""
                    # Let user pick final values
                    new_start_date = st.text_input("Start Date", value=min_start_str, key=f"{proj}_{mon}_start")
                    new_comp_date = st.text_input("Completion Date", value=max_comp_str, key=f"{proj}_{mon}_comp")

                    # For each cat field, pick from allowed
                    cat_map = {}
                    for cf, opts in allowed_values_manual.items():
                        # Just pick the first row's value as the "suggestion"
                        current_val = group[cf].iloc[0] if cf in group.columns and not group[cf].dropna().empty else ""
                        if current_val in opts:
                            idx_default = opts.index(current_val)
                        else:
                            idx_default = 0
                        cat_map[cf] = st.selectbox(f"{cf}", options=opts, index=idx_default, key=f"{proj}_{mon}_{cf}")

                    update_instructions[f"{proj}||{mon}"] = {
                        "Start Date": new_start_date,
                        "Completion Date": new_comp_date,
                        **cat_map
                    }

        st.markdown("#### Final Update Instructions:")
        st.json(update_instructions)

        if st.button("Update Data"):
            # Save instructions to text file
            try:
                with open(TEMP_UPDATE_FILE, "w", encoding="utf-8") as f:
                    json.dump(update_instructions, f, indent=2)
                st.success(f"Update instructions saved to {TEMP_UPDATE_FILE}.")
            except Exception as e:
                st.error(f"Error writing to text file: {e}")
                st.stop()

            # Now read them and update Excel
            try:
                wb = load_workbook(FILE_PATH)
            except Exception as e:
                st.error(f"Error opening workbook: {e}")
                st.stop()

            # For each row in working_details, if (Main project, Month) is in instructions, update
            for i, row in working_details.iterrows():
                proj = str(row["Main project"])  # ensure string
                mon = str(row["Month"])          # ensure string
                key = f"{proj}||{mon}"
                if key in update_instructions:
                    instructions = update_instructions[key]
                    sheet_name = row["Employee"]
                    if sheet_name not in wb.sheetnames:
                        continue
                    ws = wb[sheet_name]
                    headers = {cell.value: cell.column for cell in ws[1]}
                    r_num = row["RowNumber"]
                    if r_num < 1:
                        continue
                    if "Start Date" in headers and instructions.get("Start Date", ""):
                        ws.cell(row=r_num, column=headers["Start Date"], value=instructions["Start Date"])
                    if "Completion Date" in headers and instructions.get("Completion Date", ""):
                        ws.cell(row=r_num, column=headers["Completion Date"], value=instructions["Completion Date"])
                    # For each categorical col in allowed_values
                    for cf in allowed_values.keys():
                        if cf in headers and instructions.get(cf, ""):
                            ws.cell(row=r_num, column=headers[cf], value=instructions[cf])

            try:
                wb.save(FILE_PATH)
                st.success("Excel file updated successfully.")
            except Exception as e:
                st.error(f"Error saving workbook: {e}")
    else:
        st.info("Apply filters above to see update data.")