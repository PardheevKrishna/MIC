import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO
from openpyxl import load_workbook

# ---------- CONFIGURATION ----------
FILE_PATH = "input.xlsx"  # CHANGE this to your actual Excel file path

st.set_page_config(page_title="Team Report Dashboard", layout="wide")
st.title("Team Report Dashboard (Fixed Excel File)")

# ===================== PROCESS FUNCTION =====================
def process_excel_file(file_path):
    """
    Reads each employee sheet (employee names from "Home") and returns two DataFrames:
      - working_details: all employee data with added columns:
            Employee, RowNumber, Month (yyyy-mm), WeekFriday (mm-dd-yyyy)
      - violations_df: rows flagged as violations (Invalid value, Start date change, Weekly hours < 40)
         Each violation row gets a UniqueID (constructed later as "Employee_RowNumber")
    """
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

    # Allowed categorical values
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
    start_date_exceptions = ["Annual Leave"]

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
        df["RowNumber"] = df.index + 2  # Excel row number (header is row 1)
        df["Status Date (Every Friday)"] = pd.to_datetime(
            df["Status Date (Every Friday)"], format="%m-%d-%Y", errors="coerce"
        )

        # (1) Validate allowed values
        for col, allowed_list in allowed_values.items():
            for i, val in df[col].items():
                if pd.isna(val):
                    continue
                tokens = [t.strip() for t in str(val).split(",") if t.strip()]
                if len(tokens) != 1 or tokens[0] not in allowed_list:
                    viol_list.append({
                        "Employee": emp,
                        "Violation Type": "Invalid value",
                        "Violation Details": f"{col} = {val}",
                        "Location": f"Sheet {emp}, Row {df.at[i, 'RowNumber']}",
                        "Violation Date": df.at[i, "Status Date (Every Friday)"]
                    })

        # (2) Check start date consistency within project & month
        for i, row in df.iterrows():
            proj = row["Name of the Project"]
            start_val = row["Start Date"]
            if str(row["Main project"]).strip() in start_date_exceptions or str(proj).strip() in start_date_exceptions:
                continue
            if pd.notna(proj) and pd.notna(start_val) and pd.notna(row["Status Date (Every Friday)"]):
                month_key = row["Status Date (Every Friday)"].strftime("%Y-%m") if pd.notna(row["Status Date (Every Friday)"]) else "N/A"
                key = (proj, month_key)
                current_start = pd.to_datetime(start_val, format="%m-%d-%Y", errors="coerce")
                if key not in project_month_info:
                    project_month_info[key] = current_start
                else:
                    baseline = project_month_info[key]
                    if pd.notna(current_start) and pd.notna(baseline) and current_start != baseline:
                        old_str = baseline.strftime("%m-%d-%Y") if pd.notna(baseline) else "N/A"
                        new_str = current_start.strftime("%m-%d-%Y") if pd.notna(current_start) else "N/A"
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
            friday_str = friday.strftime("%m-%d-%Y") if pd.notna(friday) else "N/A"
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

        # (4) Add extra columns
        df["PTO Hours"] = df.apply(lambda r: r["Weekly Time Spent(Hrs)"] if "PTO" in str(r["Main project"]) else 0, axis=1)
        df["Work Hours"] = df.apply(lambda r: r["Weekly Time Spent(Hrs)"] if "PTO" not in str(r["Main project"]) else 0, axis=1)
        df["Month"] = df["Status Date (Every Friday)"].dt.to_period("M").astype(str)
        df["WeekFriday"] = df["Status Date (Every Friday)"].dt.strftime("%m-%d-%Y").fillna("N/A")
        working_list.append(df)

    if working_list:
        working_details = pd.concat(working_list, ignore_index=True)
    else:
        working_details = pd.DataFrame()

    violations_df = pd.DataFrame(viol_list)
    return working_details, violations_df

# ---------- READ DATA ----------
working_details, violations_df = process_excel_file(FILE_PATH)
if working_details is None or violations_df is None:
    st.error("Error processing the Excel file.")
    st.stop()
else:
    st.success("Reports generated successfully!")

# ========== CREATE TABS ==========
tab1, tab2, tab3 = st.tabs(["Team Monthly Summary", "Working Hours Summary", "Violations and Update"])

# ========== TAB 1: TEAM MONTHLY SUMMARY ==========
with tab1:
    st.subheader("Team Monthly Summary")
    if working_details.empty:
        st.info("No data available.")
    else:
        all_emps = sorted(working_details["Employee"].dropna().unique())
        all_months = sorted(working_details["Month"].dropna().unique())
        all_weeks = sorted(working_details["WeekFriday"].dropna().unique())
        with st.form("tm_form"):
            c1, c2 = st.columns([0.7, 0.3])
            emp_sel = c1.multiselect("Select Employee(s)", options=all_emps)
            all_emp_check = c2.checkbox("Select All Employees", key="tm_all_emp")
            c3, c4 = st.columns([0.7, 0.3])
            month_sel = c3.multiselect("Select Month(s)", options=all_months)
            all_month_check = c4.checkbox("Select All Months", key="tm_all_month")
            if month_sel:
                possible_weeks = sorted(working_details[working_details["Month"].isin(month_sel)]["WeekFriday"].dropna().unique())
            else:
                possible_weeks = all_weeks
            c5, c6 = st.columns([0.7, 0.3])
            week_sel = c5.multiselect("Select Week(s)", options=possible_weeks)
            all_week_check = c6.checkbox("Select All Weeks", key="tm_all_week")
            submit_tm = st.form_submit_button("Filter Data")
        if submit_tm:
            if all_emp_check:
                emp_sel = all_emps
            if all_month_check:
                month_sel = all_months
            if all_week_check:
                week_sel = possible_weeks
            df_tm = working_details.copy()
            if emp_sel:
                df_tm = df_tm[df_tm["Employee"].isin(emp_sel)]
            if month_sel:
                df_tm = df_tm[df_tm["Month"].isin(month_sel)]
            if week_sel:
                df_tm = df_tm[df_tm["WeekFriday"].isin(week_sel)]
            if month_sel:
                summary = df_tm.groupby(["Employee", "Month", "WeekFriday"]).agg({"Work Hours": "sum", "PTO Hours": "sum"}).reset_index()
            else:
                summary = df_tm.groupby(["Employee", "Month"]).agg({"Work Hours": "sum", "PTO Hours": "sum"}).reset_index()
            st.dataframe(summary, use_container_width=True)
            buf = BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                summary.to_excel(writer, sheet_name="Team_Monthly", index=False)
            buf.seek(0)
            st.download_button("Download Team Monthly Summary", data=buf, file_name="Team_Monthly_Summary.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ========== TAB 2: WORKING HOURS SUMMARY ==========
with tab2:
    st.subheader("Working Hours Summary")
    if working_details.empty:
        st.info("No data available.")
    else:
        all_emps_wh = sorted(working_details["Employee"].dropna().unique())
        all_months_wh = sorted(working_details["Month"].dropna().unique())
        all_weeks_wh = sorted(working_details["WeekFriday"].dropna().unique())
        with st.form("wh_form"):
            col1, col2 = st.columns([0.7, 0.3])
            emp_sel_wh = col1.multiselect("Select Employee(s)", options=all_emps_wh)
            all_emp_wh = col2.checkbox("Select All Employees", key="wh_all_emp")
            col3, col4 = st.columns([0.7, 0.3])
            month_sel_wh = col3.multiselect("Select Month(s)", options=all_months_wh)
            all_month_wh = col4.checkbox("Select All Months", key="wh_all_month")
            if month_sel_wh:
                poss_weeks_wh = sorted(working_details[working_details["Month"].isin(month_sel_wh)]["WeekFriday"].dropna().unique())
            else:
                poss_weeks_wh = all_weeks_wh
            col5, col6 = st.columns([0.7, 0.3])
            week_sel_wh = col5.multiselect("Select Week(s)", options=poss_weeks_wh)
            all_week_wh = col6.checkbox("Select All Weeks", key="wh_all_week")
            submit_wh = st.form_submit_button("Filter Data")
        if submit_wh:
            if all_emp_wh:
                emp_sel_wh = all_emps_wh
            if all_month_wh:
                month_sel_wh = all_months_wh
            if all_week_wh:
                week_sel_wh = poss_weeks_wh
            df_wh = working_details.copy()
            if emp_sel_wh:
                df_wh = df_wh[df_wh["Employee"].isin(emp_sel_wh)]
            if month_sel_wh:
                df_wh = df_wh[df_wh["Month"].isin(month_sel_wh)]
            if week_sel_wh:
                df_wh = df_wh[df_wh["WeekFriday"].isin(week_sel_wh)]
            st.dataframe(df_wh, use_container_width=True)
            buf_wh = BytesIO()
            with pd.ExcelWriter(buf_wh, engine="openpyxl") as writer:
                df_wh.to_excel(writer, sheet_name="Working_Hours", index=False)
            buf_wh.seek(0)
            st.download_button("Download Working Hours", data=buf_wh, file_name="Working_Hours_Summary.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ========== TAB 3: VIOLATIONS & UPDATE ==========
with tab3:
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
            # Create UniqueID as "Employee_RowNumber" from the Location field
            def extract_rownum(loc_str):
                try:
                    return loc_str.split("Row ")[-1]
                except:
                    return ""
            df_v["UniqueID"] = df_v.apply(lambda r: f"{r['Employee']}_{extract_rownum(r['Location'])}", axis=1)
            st.dataframe(df_v, use_container_width=True)

            # Step 2: Select rows to update
            st.markdown("#### Select Rows to Update (by UniqueID)")
            all_ids = sorted(df_v["UniqueID"].unique())
            select_all_rows = st.checkbox("Select All Rows")
            if select_all_rows:
                selected_ids = all_ids
            else:
                selected_ids = st.multiselect("Select UniqueIDs", options=all_ids)

            # Step 3: Proceed to Update (load row-edit forms)
            if st.button("Proceed to Update"):
                if not selected_ids:
                    st.error("No rows selected for update.")
                else:
                    st.session_state["selected_rows"] = selected_ids
                    st.markdown(f"**Rows selected for update:** {selected_ids}")
                    # Choose update mode now:
                    upd_mode = st.radio("Select Update Mode", options=["Automatic", "Manual"], index=0)
                    
                    st.markdown("### Edit Details for Each Selected Row")
                    # Prepare to collect updated data
                    updated_rows = {}
                    # For each selected UniqueID, display an expander with a form
                    for uid in selected_ids:
                        row_data = working_details[working_details["UniqueID"] == uid]
                        if row_data.empty:
                            continue
                        row = row_data.iloc[0]
                        with st.expander(f"Edit details for {uid} (Employee: {row['Employee']})", expanded=True):
                            # For Automatic mode, compute suggestions within this row's main project group
                            if upd_mode == "Automatic":
                                group = working_details[working_details["Main project"] == row["Main project"]]
                                sug_start = group["Start Date"].apply(lambda x: pd.to_datetime(x, errors="coerce")).min()
                                sug_comp = (group["Completion Date"].apply(lambda x: pd.to_datetime(x, errors="coerce")).max()
                                            if "Completion Date" in group.columns else None)
                                # For categorical fields, use most frequent value in the group
                                sug_values = {}
                                for field in ["Functional Area (CRIT, CRIT - Data Management, CRIT - Data Governance, CRIT - Regulatory Reporting, CRIT - Portfolio Reporting, CRIT - Transformation)",
                                              "Project Category (Data Infrastructure, Monitoring & Insights, Analytics / Strategy Development, GDA Related, Trainings and Team Meeting)",
                                              "Complexity (H,M,L)",
                                              "Novelity (BAU repetitive, One time repetitive, New one time)",
                                              "Output Type (Core production work, Ad-hoc long-term projects, Ad-hoc short-term projects, Business Management, Administration, Trainings/L&D activities, Others) :",
                                              "Impact type (Customer Experience, Financial impact, Insights, Risk reduction, Others)"]:
                                    if field in group.columns and not group[field].dropna().empty:
                                        sug_values[field] = group[field].mode().iloc[0]
                                    else:
                                        sug_values[field] = ""
                                # Convert suggested dates to strings safely
                                if pd.notna(sug_start):
                                    sug_start_str = sug_start.strftime("%m-%d-%Y")
                                else:
                                    sug_start_str = ""
                                if sug_comp is not None and pd.notna(sug_comp):
                                    sug_comp_str = sug_comp.strftime("%m-%d-%Y")
                                else:
                                    sug_comp_str = ""
                            else:
                                # Manual mode: use current row values
                                current_start = pd.to_datetime(row["Start Date"], errors="coerce")
                                if pd.notna(current_start):
                                    sug_start_str = current_start.strftime("%m-%d-%Y")
                                else:
                                    sug_start_str = ""
                                if "Completion Date" in row and pd.notna(row.get("Completion Date", None)):
                                    current_comp = pd.to_datetime(row["Completion Date"], errors="coerce")
                                    if pd.notna(current_comp):
                                        sug_comp_str = current_comp.strftime("%m-%d-%Y")
                                    else:
                                        sug_comp_str = ""
                                else:
                                    sug_comp_str = ""
                                sug_values = {}
                                for field in ["Functional Area (CRIT, CRIT - Data Management, CRIT - Data Governance, CRIT - Regulatory Reporting, CRIT - Portfolio Reporting, CRIT - Transformation)",
                                              "Project Category (Data Infrastructure, Monitoring & Insights, Analytics / Strategy Development, GDA Related, Trainings and Team Meeting)",
                                              "Complexity (H,M,L)",
                                              "Novelity (BAU repetitive, One time repetitive, New one time)",
                                              "Output Type (Core production work, Ad-hoc long-term projects, Ad-hoc short-term projects, Business Management, Administration, Trainings/L&D activities, Others) :",
                                              "Impact type (Customer Experience, Financial impact, Insights, Risk reduction, Others)"]:
                                    sug_values[field] = row.get(field, "")

                            # Now, show editable fields for this row
                            new_start = st.date_input("Start Date", value=datetime.strptime(sug_start_str, "%m-%d-%Y") if sug_start_str else datetime.today(), key=f"{uid}_start")
                            if "Completion Date" in row:
                                new_comp = st.date_input("Completion Date", value=datetime.strptime(sug_comp_str, "%m-%d-%Y") if sug_comp_str else datetime.today(), key=f"{uid}_comp")
                            else:
                                new_comp = None
                            new_data = {
                                "Start Date": new_start.strftime("%m-%d-%Y"),
                                "Completion Date": new_comp.strftime("%m-%d-%Y") if new_comp else ""
                            }
                            for field in categorical_fields:
                                # For automatic mode, you might want to use a selectbox pre-filled with suggestion and allowed options.
                                # We'll use allowed options from our allowed_values dictionary.
                                allowed_opts = allowed_values.get(field, [])
                                new_val = st.selectbox(f"{field}", options=allowed_opts, index=allowed_opts.index(sug_values[field]) if sug_values[field] in allowed_opts else 0, key=f"{uid}_{field}")
                                new_data[field] = new_val
                            updated_rows[uid] = {
                                "Employee": row["Employee"],
                                "RowNumber": row["RowNumber"],
                                **new_data
                            }
                    # End of per-row editing loop
                    if st.button("Update Details"):
                        # Now update the Excel file using openpyxl
                        try:
                            wb = load_workbook(FILE_PATH)
                        except Exception as e:
                            st.error(f"Error opening workbook: {e}")
                            st.stop()
                        for uid, new_vals in updated_rows.items():
                            sheet_name = new_vals["Employee"]  # assume sheet name equals employee name
                            if sheet_name not in wb.sheetnames:
                                continue
                            ws = wb[sheet_name]
                            # Get header mapping (assume header in row 1)
                            headers = {cell.value: cell.column for cell in ws[1]}
                            r_num = new_vals["RowNumber"]
                            if "Start Date" in headers:
                                ws.cell(row=r_num, column=headers["Start Date"], value=new_vals["Start Date"])
                            if "Completion Date" in headers and new_vals.get("Completion Date", ""):
                                ws.cell(row=r_num, column=headers["Completion Date"], value=new_vals["Completion Date"])
                            for field in categorical_fields:
                                if field in headers:
                                    ws.cell(row=r_num, column=headers[field], value=new_vals[field])
                        try:
                            wb.save(FILE_PATH)
                            st.success("Excel file updated successfully.")
                        except Exception as e:
                            st.error(f"Error saving workbook: {e}")
        else:
            st.info("Please use the form above to filter violations.")