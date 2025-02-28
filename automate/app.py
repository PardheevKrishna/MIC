import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO
from openpyxl import load_workbook

# ---------- CONFIGURATION ----------
FILE_PATH = "input.xlsx"  # Change this to your actual Excel file path

st.set_page_config(page_title="Team Report Dashboard", layout="wide")
st.title("Team Report Dashboard (Fixed Excel File)")

# ===================== PROCESS FUNCTION =====================
def process_excel_file(file_path):
    """
    Reads each employee sheet (employee names are taken from the "Home" sheet)
    and returns two DataFrames:
      - working_details: all employee data with added columns:
            Employee, RowNumber, Month (yyyy-mm), WeekFriday (mm-dd-yyyy)
      - violations_df: rows flagged as violations, with columns:
            Employee, Violation Type, Violation Details, Location, Violation Date
         (Violation Type is one of: "Invalid value", "Working hours less than 40", "Start date change")
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
        if not all(col in df.columns for col in req_cols):
            continue
        df["Employee"] = emp
        df["RowNumber"] = df.index + 2  # header is row 1
        df["Status Date (Every Friday)"] = pd.to_datetime(df["Status Date (Every Friday)"], format="%m-%d-%Y", errors="coerce")
        
        # (1) Allowed values validation
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
        # (2) Start Date consistency check
        for i, row in df.iterrows():
            proj = row["Name of the Project"]
            start_val = row["Start Date"]
            if str(row["Main project"]).strip() in start_date_exceptions or str(proj).strip() in start_date_exceptions:
                continue
            if pd.notna(proj) and pd.notna(start_val) and pd.notna(row["Status Date (Every Friday)"]):
                month_key = row["Status Date (Every Friday)"].strftime("%Y-%m")
                key = (proj, month_key)
                current_start = pd.to_datetime(start_val, format="%m-%d-%Y", errors="coerce")
                if key not in project_month_info:
                    project_month_info[key] = current_start
                else:
                    if current_start != project_month_info[key]:
                        viol_list.append({
                            "Employee": emp,
                            "Violation Type": "Start date change",
                            "Violation Details": f"{proj}: expected {project_month_info[key].strftime('%m-%d-%Y')}, got {current_start.strftime('%m-%d-%Y')}",
                            "Location": f"Sheet {emp}, Row {row['RowNumber']}",
                            "Violation Date": row["Status Date (Every Friday)"]
                        })
        # (3) Weekly hours check
        df["Weekly Time Spent(Hrs)"] = pd.to_numeric(df["Weekly Time Spent(Hrs)"], errors="coerce").fillna(0)
        for friday in df[df["Status Date (Every Friday)"].dt.weekday == 4]["Status Date (Every Friday)"].unique():
            week_start = friday - timedelta(days=4)
            week_df = df[(df["Status Date (Every Friday)"] >= week_start) & (df["Status Date (Every Friday)"] <= friday)]
            if week_df["Weekly Time Spent(Hrs)"].sum() < 40:
                viol_list.append({
                    "Employee": emp,
                    "Violation Type": "Working hours less than 40",
                    "Violation Details": f"Week ending {friday.strftime('%m-%d-%Y')} insufficient hours",
                    "Location": f"Sheet {emp}, Rows: {', '.join(week_df['RowNumber'].astype(str))}",
                    "Violation Date": friday
                })
        # (4) Additional columns
        df["PTO Hours"] = df.apply(lambda r: r["Weekly Time Spent(Hrs)"] if "PTO" in str(r["Main project"]) else 0, axis=1)
        df["Work Hours"] = df.apply(lambda r: r["Weekly Time Spent(Hrs)"] if "PTO" not in str(r["Main project"]) else 0, axis=1)
        df["Month"] = df["Status Date (Every Friday)"].dt.to_period("M").astype(str)
        df["WeekFriday"] = df["Status Date (Every Friday)"].dt.strftime("%m-%d-%Y")
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
    st.error("Error processing Excel file.")
else:
    st.success("Reports generated successfully!")
    
    # ---------- TEAM MONTHLY SUMMARY Tab ----------
    tab1, tab2, tab3 = st.tabs(["Team Monthly Summary", "Working Hours Summary", "Violations and Update"])
    
    with tab1:
        st.subheader("Team Monthly Summary")
        if working_details.empty:
            st.info("No data available.")
        else:
            all_emps = sorted(working_details["Employee"].dropna().unique())
            all_months = sorted(working_details["Month"].dropna().unique())
            all_weeks = sorted(working_details["WeekFriday"].dropna().unique())
            with st.form("tm_form"):
                col_a, col_b = st.columns([0.7, 0.3])
                emp_sel = col_a.multiselect("Select Employee(s)", options=all_emps)
                all_emp_check = col_b.checkbox("Select All Employees", key="tm_all_emp")
                col_c, col_d = st.columns([0.7, 0.3])
                month_sel = col_c.multiselect("Select Month(s)", options=all_months)
                all_month_check = col_d.checkbox("Select All Months", key="tm_all_month")
                if month_sel:
                    possible_weeks = sorted(working_details[working_details["Month"].isin(month_sel)]["WeekFriday"].dropna().unique())
                else:
                    possible_weeks = all_weeks
                col_e, col_f = st.columns([0.7, 0.3])
                week_sel = col_e.multiselect("Select Week(s)", options=possible_weeks)
                all_week_check = col_f.checkbox("Select All Weeks", key="tm_all_week")
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
    
    # ---------- WORKING HOURS SUMMARY Tab ----------
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
    
    # ---------- VIOLATIONS and UPDATE Tab ----------
    with tab3:
        st.subheader("Violations and Update")
        if violations_df.empty:
            st.info("No violations found.")
        else:
            all_emps_v = sorted(violations_df["Employee"].dropna().unique())
            all_types_v = ["Invalid value", "Working hours less than 40", "Start date change"]
            emp_sel_v = st.multiselect("Select Employee(s)", options=all_emps_v)
            type_sel_v = st.multiselect("Select Violation Type(s)", options=all_types_v)
            df_v = violations_df.copy()
            if emp_sel_v:
                df_v = df_v[df_v["Employee"].isin(emp_sel_v)]
            if type_sel_v:
                df_v = df_v[df_v["Violation Type"].isin(type_sel_v)]
            # Create UniqueID for each violation from Employee and row number extracted from Location.
            def extract_row(loc_str):
                try:
                    return loc_str.split("Row ")[-1]
                except:
                    return ""
            df_v["UniqueID"] = df_v.apply(lambda r: f"{r['Employee']}_{extract_row(r['Location'])}", axis=1)
            st.dataframe(df_v, use_container_width=True)
            st.markdown("#### Select Rows to Update (by UniqueID)")
            selected_ids = st.multiselect("Select UniqueIDs", options=df_v["UniqueID"].unique())
            
            if st.button("Proceed to Update"):
                if not selected_ids:
                    st.error("Please select at least one row for update.")
                else:
                    st.session_state["selected_rows"] = selected_ids
                    st.markdown(f"**Rows selected for update:** {selected_ids}")
                    
                    st.markdown("### Update Options")
                    upd_mode = st.radio("Select Update Mode", options=["Automatic", "Manual"], index=0)
                    categorical_fields = [
                        "Functional Area (CRIT, CRIT - Data Management, CRIT - Data Governance, CRIT - Regulatory Reporting, CRIT - Portfolio Reporting, CRIT - Transformation)",
                        "Project Category (Data Infrastructure, Monitoring & Insights, Analytics / Strategy Development, GDA Related, Trainings and Team Meeting)",
                        "Complexity (H,M,L)",
                        "Novelity (BAU repetitive, One time repetitive, New one time)",
                        "Output Type (Core production work, Ad-hoc long-term projects, Ad-hoc short-term projects, Business Management, Administration, Trainings/L&D activities, Others) :",
                        "Impact type (Customer Experience, Financial impact, Insights, Risk reduction, Others)"
                    ]
                    # Re-read full working details to compute suggestions for the selected rows.
                    working_details["UniqueID"] = working_details.apply(lambda r: f"{r['Employee']}_{r['RowNumber']}", axis=1)
                    sel_rows = working_details[working_details["UniqueID"].isin(selected_ids)]
                    st.markdown("#### Selected Rows Preview")
                    st.dataframe(sel_rows, use_container_width=True)
                    update_options = {}
                    if upd_mode == "Automatic":
                        st.markdown("**Automatic Mode**")
                        if not sel_rows.empty:
                            auto_start = sel_rows["Start Date"].apply(lambda x: pd.to_datetime(x, format="%m-%d-%Y", errors="coerce")).min()
                            if "Completion Date" in sel_rows.columns:
                                auto_comp = sel_rows["Completion Date"].apply(lambda x: pd.to_datetime(x, format="%m-%d-%Y", errors="coerce")).max()
                            else:
                                auto_comp = None
                        else:
                            st.error("No rows selected for update.")
                            auto_start, auto_comp = None, None
                        st.write("Auto Start Date:", auto_start.strftime("%m-%d-%Y") if pd.notna(auto_start) else "N/A")
                        st.write("Auto Completion Date:", auto_comp.strftime("%m-%d-%Y") if auto_comp is not None and pd.notna(auto_comp) else "N/A")
                        auto_choices = {}
                        auto_sugg = {}
                        for field in categorical_fields:
                            if field in sel_rows.columns and not sel_rows[field].dropna().empty:
                                first_occ = sel_rows[field].iloc[0]
                                most_freq = sel_rows[field].mode().iloc[0]
                            else:
                                first_occ, most_freq = None, None
                            auto_sugg[field] = {"First Occurrence": first_occ, "Most Frequent": most_freq}
                            auto_choices[field] = st.radio(f"Update {field} with", options=["First Occurrence", "Most Frequent"], index=0, key=field+"_upd_auto")
                        update_options["Start Date"] = auto_start.strftime("%m-%d-%Y") if pd.notna(auto_start) else None
                        update_options["Completion Date"] = auto_comp.strftime("%m-%d-%Y") if auto_comp is not None and pd.notna(auto_comp) else None
                        for field in categorical_fields:
                            update_options[field] = auto_sugg[field][auto_choices[field]]
                    else:
                        st.markdown("**Manual Mode**")
                        manual_start = st.date_input("Select Start Date", value=datetime.today())
                        manual_comp = st.date_input("Select Completion Date", value=datetime.today())
                        update_options["Start Date"] = manual_start.strftime("%m-%d-%Y")
                        update_options["Completion Date"] = manual_comp.strftime("%m-%d-%Y")
                        manual_vals = {}
                        allowed_manual = {
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
                        for field in categorical_fields:
                            manual_vals[field] = st.selectbox(f"Select value for {field}", options=allowed_manual[field], key=field+"_upd_manual")
                        update_options.update(manual_vals)
                    st.markdown("#### Final Update Options")
                    st.write(update_options)
                    if st.button("Update Excel"):
                        try:
                            wb = load_workbook(FILE_PATH)
                        except Exception as e:
                            st.error(f"Error opening workbook: {e}")
                        for sheet in wb.sheetnames:
                            if sheet == "Home":
                                continue
                            ws = wb[sheet]
                            # Assume header is in row 1; map header names to columns.
                            headers = {cell.value: cell.column for cell in ws[1]}
                            if "Employee" not in headers or "Start Date" not in headers:
                                continue
                            for r in range(2, ws.max_row + 1):
                                emp_val = ws.cell(row=r, column=headers["Employee"]).value
                                unique_id = f"{emp_val}_{r}"
                                if unique_id in st.session_state["selected_rows"]:
                                    if "Start Date" in headers and update_options.get("Start Date"):
                                        ws.cell(row=r, column=headers["Start Date"], value=update_options["Start Date"])
                                    if "Completion Date" in headers and update_options.get("Completion Date"):
                                        ws.cell(row=r, column=headers["Completion Date"], value=update_options["Completion Date"])
                                    for field in categorical_fields:
                                        if field in headers and update_options.get(field):
                                            ws.cell(row=r, column=headers[field], value=update_options[field])
                        try:
                            wb.save(FILE_PATH)
                            st.success("Excel file updated successfully.")
                        except Exception as e:
                            st.error(f"Error saving workbook: {e}")
        else:
            st.info("Please filter violations and then select rows to update.")