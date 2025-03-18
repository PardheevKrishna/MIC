import streamlit as st
st.set_page_config(page_title="Final FRY14M Field Analysis", layout="wide", initial_sidebar_state="expanded")

import pandas as pd
import os
import datetime
import re
from io import BytesIO
from dateutil.relativedelta import relativedelta
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode
from openpyxl import load_workbook
from openpyxl.comments import Comment

#############################################
# Helper Functions
#############################################

def compute_grid_height(df, row_height=40, header_height=80):
    n = len(df)
    return header_height + (min(n, 30) * row_height)

def get_excel_engine(file_path):
    return 'pyxlsb' if file_path.lower().endswith('.xlsb') else None

def generate_summary_df(df_data, date1, date2):
    fields = sorted(df_data["field_name"].unique())
    rows = []
    for field in fields:
        mask_miss_d1 = ((df_data['analysis_type'] == 'value_dist') &
                        (df_data['field_name'] == field) &
                        (df_data['filemonth_dt'] == date1) &
                        (df_data['value_label'].str.contains("Missing", case=False, na=False)))
        missing_d1 = df_data.loc[mask_miss_d1, 'value_records'].sum()
        mask_miss_d2 = ((df_data['analysis_type'] == 'value_dist') &
                        (df_data['field_name'] == field) &
                        (df_data['filemonth_dt'] == date2) &
                        (df_data['value_label'].str.contains("Missing", case=False, na=False)))
        missing_d2 = df_data.loc[mask_miss_d2, 'value_records'].sum()
        phrases = [
            "1\\)   CF Loan - Both Pop, Diff Values",
            "2\\)   CF Loan - Prior Null, Current Pop",
            "3\\)   CF Loan - Prior Pop, Current Null"
        ]
        def contains_phrase(x):
            for pat in phrases:
                if pd.notna(x) and re.search(pat, x):
                    return True
            return False
        mask_pop_d1 = ((df_data['analysis_type'] == 'pop_comp') &
                       (df_data['field_name'] == field) &
                       (df_data['filemonth_dt'] == date1) &
                       (df_data['value_label'].apply(contains_phrase)))
        pop_d1 = df_data.loc[mask_pop_d1, 'value_records'].sum()
        mask_pop_d2 = ((df_data['analysis_type'] == 'pop_comp') &
                       (df_data['field_name'] == field) &
                       (df_data['filemonth_dt'] == date2) &
                       (df_data['value_label'].apply(contains_phrase)))
        pop_d2 = df_data.loc[mask_pop_d2, 'value_records'].sum()
        rows.append([field, missing_d1, missing_d2, pop_d1, pop_d2])
    df = pd.DataFrame(rows, columns=[
        "Field Name",
        f"Missing Values ({date1.strftime('%Y-%m-%d')})",
        f"Missing Values ({date2.strftime('%Y-%m-%d')})",
        f"Month-to-Month Diff ({date1.strftime('%Y-%m-%d')})",
        f"Month-to-Month Diff ({date2.strftime('%Y-%m-%d')})"
    ])
    m1 = f"Missing Values ({date1.strftime('%Y-%m-%d')})"
    m2 = f"Missing Values ({date2.strftime('%Y-%m-%d')})"
    d1 = f"Month-to-Month Diff ({date1.strftime('%Y-%m-%d')})"
    d2 = f"Month-to-Month Diff ({date2.strftime('%Y-%m-%d')})"
    df["Missing % Change"] = df.apply(lambda r: ((r[m1]-r[m2]) / r[m2] * 100) if r[m2]!=0 else None, axis=1)
    df["Month-to-Month % Change"] = df.apply(lambda r: ((r[d1]-r[d2]) / r[d2] * 100) if r[d2]!=0 else None, axis=1)
    new_order = [ "Field Name", m1, m2, "Missing % Change", d1, d2, "Month-to-Month % Change" ]
    df = df[new_order]
    # Add two extra columns: one for current aggregated comments and one for editable Approval Comments.
    df["Comment"] = ""
    df["Approval Comments"] = ""
    return df

def generate_distribution_df(df, analysis_type, date1):
    months = [(date1 - relativedelta(months=i)).replace(day=1) for i in range(12)]
    months = sorted(months, reverse=True)
    sub = df[df['analysis_type'] == analysis_type].copy()
    sub['month'] = sub['filemonth_dt'].apply(lambda d: d.replace(day=1))
    sub = sub[sub['month'].isin(months)]
    grouped = sub.groupby(['field_name', 'value_label', 'month'])['value_records'].sum().reset_index()
    if grouped.empty:
        return pd.DataFrame()
    pivot = grouped.pivot_table(index=['field_name', 'value_label'],
                                columns='month', values='value_records', fill_value=0)
    pivot = pivot.reindex(columns=months, fill_value=0)
    frames = []
    for field, sub_df in pivot.groupby(level=0):
        sub_df = sub_df.droplevel(0)
        total = sub_df.sum(axis=0)
        pct = sub_df.div(total, axis=1).mul(100).round(2).fillna(0)
        data = {}
        for m in months:
            m_str = m.strftime('%Y-%m')
            data[(m_str, "Sum")] = sub_df[m]
            data[(m_str, "Percent")] = pct[m]
        tmp = pd.DataFrame(data)
        tot_row = {}
        for m in months:
            m_str = m.strftime('%Y-%m')
            tot_row[(m_str, "Sum")] = total[m]
            tot_row[(m_str, "Percent")] = ""
        tmp.loc["Current period total"] = tot_row
        tmp.index = pd.MultiIndex.from_product([[field], tmp.index], names=["Field Name", "Value Label"])
        frames.append(tmp)
    if not frames:
        return pd.DataFrame()
    final = pd.concat(frames)
    final.columns = pd.MultiIndex.from_tuples(final.columns)
    return final

def flatten_dataframe(df):
    if isinstance(df.columns, pd.MultiIndex):
        df = df.reset_index()
        df.columns = [" ".join(map(str, col)).strip() if isinstance(col, tuple) else col for col in df.columns.values]
    return df

def load_report_data(file_path, date1, date2):
    engine = get_excel_engine(file_path)
    if engine:
        df_data = pd.read_excel(file_path, sheet_name="Data", engine=engine)
    else:
        df_data = pd.read_excel(file_path, sheet_name="Data")
    df_data["filemonth_dt"] = pd.to_datetime(df_data["filemonth_dt"])
    summary_df = generate_summary_df(df_data, date1, date2)
    val_dist_df = generate_distribution_df(df_data, "value_dist", date1)
    pop_comp_df = generate_distribution_df(df_data, "pop_comp", date1)
    return df_data, summary_df, val_dist_df, pop_comp_df

#############################################
# Functions to Cache Previous Comments
#############################################

def cache_previous_comments(current_folder):
    """
    Scans "previous/<current_folder>" for .xlsx files.
    Extracts the month-year from the filename (e.g. "202412" or "2024-12")
    and uses that as the Month. Then extracts cell comments from the "Summary" sheet
    (using a header that contains "month to month") and writes a CSV with columns: Field Name, Month, Comment.
    """
    data = []
    prev_folder = os.path.join(os.getcwd(), "previous", current_folder)
    if not os.path.exists(prev_folder):
        st.warning("Previous months folder not found.")
        return pd.DataFrame(columns=["Field Name", "Month", "Comment"])
    month_pattern = re.compile(r'(\d{4})[- ]?(\d{2})')
    for file in os.listdir(prev_folder):
        if file.lower().endswith('.xlsx'):
            file_path = os.path.join(prev_folder, file)
            m = month_pattern.search(file)
            if m:
                month_year = f"{m.group(1)}-{m.group(2)}"
            else:
                month_year = "unknown"
            try:
                wb = load_workbook(file_path, data_only=True)
            except Exception as e:
                st.error(f"Error opening previous file {file}: {e}")
                continue
            if "Summary" not in wb.sheetnames:
                continue
            ws = wb["Summary"]
            header_row = None
            for r in range(1, 4):
                headers = [cell.value for cell in ws[r] if cell.value is not None]
                if any("month to month" in str(val).lower() for val in headers):
                    header_row = r
                    break
            if header_row is None:
                continue
            header = [cell.value for cell in ws[header_row]]
            col_index = None
            for i, col_name in enumerate(header, start=1):
                if col_name and "month to month" in str(col_name).lower():
                    col_index = i
                    break
            if col_index is None:
                continue
            for row in ws.iter_rows(min_row=header_row+1):
                field_cell = row[0]  # Assume "Field Name" is in the first column
                if field_cell.value:
                    field_name = str(field_cell.value).strip()
                    cell = row[col_index - 1]  # 0-indexed
                    comment_text = cell.comment.text if (cell.comment and cell.comment.text is not None) else ""
                    data.append({"Field Name": field_name, "Month": month_year, "Comment": comment_text})
    df = pd.DataFrame(data)
    df.to_csv("previous_comments.csv", index=False)
    return df

def get_cached_previous_comments(current_folder):
    try:
        df = pd.read_csv("previous_comments.csv")
        if df.empty:
            return pd.DataFrame(columns=["Field Name", "Month", "Comment"])
        return df
    except pd.errors.EmptyDataError:
        return pd.DataFrame(columns=["Field Name", "Month", "Comment"])

def pivot_previous_comments(df, target_month):
    if df.empty:
        return pd.DataFrame()
    df_target = df[df["Month"] == target_month]
    if df_target.empty:
        return pd.DataFrame()
    grouped = df_target.groupby(["Field Name"])["Comment"].apply(
        lambda x: "\n".join([str(item) for item in x if pd.notnull(item)])
    ).reset_index()
    grouped = grouped.rename(columns={"Comment": f"comment_{target_month}"})
    return grouped

#############################################
# Function to Preserve Existing Summary Comments
#############################################

def preserve_summary_comments(input_file_path, summary_df):
    try:
        wb = load_workbook(input_file_path, data_only=True)
        if "Summary" in wb.sheetnames:
            existing = pd.read_excel(input_file_path, sheet_name="Summary")
            if "Approval Comments" in existing.columns:
                summary_df = summary_df.merge(existing[["Field Name", "Approval Comments"]], on="Field Name", how="left", suffixes=("", "_old"))
                summary_df["Approval Comments"] = summary_df.apply(lambda r: r["Approval Comments_old"] if pd.notnull(r["Approval Comments_old"]) else r["Approval Comments"], axis=1)
                summary_df.drop(columns=["Approval Comments_old"], inplace=True)
            if "Comment" in existing.columns:
                summary_df = summary_df.merge(existing[["Field Name", "Comment"]], on="Field Name", how="left", suffixes=("", "_old"))
                summary_df["Comment"] = summary_df.apply(lambda r: r["Comment_old"] if pd.notnull(r["Comment_old"]) and r["Comment_old"] != "" else r["Comment"], axis=1)
                summary_df.drop(columns=["Comment_old"], inplace=True)
    except Exception as e:
        st.warning(f"Could not preserve existing summary comments: {e}")
    return summary_df

#############################################
# Main Streamlit App
#############################################

def main():
    st.sidebar.title("File & Date Selection")
    folder = st.sidebar.selectbox("Select Folder", ["BDCOM", "WFHMSA", "BCards"])
    folder_path = os.path.join(os.getcwd(), folder)
    st.sidebar.write(f"Folder path: {folder_path}")
    if not os.path.exists(folder_path):
        st.sidebar.error(f"Folder '{folder}' not found.")
        return
    all_files = [f for f in os.listdir(folder_path) if f.lower().endswith(('.xlsx', '.xlsb'))]
    if not all_files:
        st.sidebar.error(f"No Excel files found in folder '{folder}'.")
        return
    selected_file = st.sidebar.selectbox("Select an Excel File", all_files)
    input_file_path = os.path.join(folder_path, selected_file)
    selected_date = st.sidebar.date_input("Select Date for Date1", datetime.date(2025, 1, 1))
    date1 = datetime.datetime.combine(selected_date, datetime.datetime.min.time())
    date2 = date1 - relativedelta(months=1)
    
    # Cache previous comments immediately.
    prev_comments_df = get_cached_previous_comments(folder)
    if prev_comments_df.empty:
        prev_comments_df = cache_previous_comments(folder)
    st.write("Cached previous comments:")
    st.dataframe(prev_comments_df)
    
    # Determine target previous month from current file's date1.
    target_prev_month = (date1 - relativedelta(months=1)).strftime("%Y-%m")
    
    if st.sidebar.button("Generate Report"):
        df_data, summary_df, val_dist_df, pop_comp_df = load_report_data(input_file_path, date1, date2)
        # Preserve existing comments from the input Excel file.
        summary_df = preserve_summary_comments(input_file_path, summary_df)
        st.session_state.df_data = df_data
        st.session_state.summary_df = summary_df
        st.session_state.value_dist_df = flatten_dataframe(val_dist_df.copy())
        st.session_state.pop_comp_df = flatten_dataframe(pop_comp_df.copy())
        st.session_state.selected_file = selected_file
        st.session_state.folder = folder
        st.session_state.date1 = date1
        st.session_state.date2 = date2
        st.session_state.input_file_path = input_file_path
        st.session_state.active_field = None

    st.write("Working Directory:", os.getcwd())
    
    if "df_data" in st.session_state:
        st.title("FRY14M Field Analysis Summary Report")
        st.write(f"**Folder:** {st.session_state.folder}")
        st.write(f"**File:** {st.session_state.selected_file}")
        st.write(f"**Date1:** {st.session_state.date1.strftime('%Y-%m-%d')} | **Date2:** {st.session_state.date2.strftime('%Y-%m-%d')}")
        
        # ---------------------------
        # Display Value Distribution Grid with "Prev Comments" Column
        st.subheader("Value Distribution")
        val_fields = st.session_state.value_dist_df["Field Name"].unique().tolist()
        if not val_fields:
            st.warning("No Value Distribution data available.")
        else:
            active_val = st.session_state.active_field if st.session_state.active_field in val_fields else val_fields[0]
            selected_val_field = st.selectbox("Select Field (Value Dist)", val_fields, index=val_fields.index(active_val), key="val_field_select")
            st.session_state.active_field = selected_val_field
            filtered_val = st.session_state.value_dist_df[st.session_state.value_dist_df["Field Name"] == selected_val_field].copy()
            if "Comment" not in filtered_val.columns:
                filtered_val["Comment"] = ""
            if "Prev Comments" in filtered_val.columns:
                filtered_val = filtered_val.drop(columns=["Prev Comments"])
            pivot_prev = pivot_previous_comments(prev_comments_df, target_prev_month)
            if not pivot_prev.empty:
                filtered_val = filtered_val.merge(pivot_prev, on="Field Name", how="left")
                filtered_val.rename(columns={f"comment_{target_prev_month}": "Prev Comments"}, inplace=True)
            else:
                filtered_val["Prev Comments"] = ""
            gb_val = GridOptionsBuilder.from_dataframe(filtered_val)
            gb_val.configure_default_column(editable=True, cellStyle={'white-space': 'normal', 'line-height': '1.2em', 'width': 150})
            gb_val.configure_column("Comment", editable=True, width=150, minWidth=100, maxWidth=200)
            gb_val.configure_column("Prev Comments", editable=False, width=150, minWidth=100, maxWidth=200)
            val_opts = gb_val.build()
            if isinstance(val_opts, list):
                val_opts = {"columnDefs": val_opts}
            val_opts["rowSelection"] = "single"
            val_opts["pagination"] = False
            val_opts["rowHeight"] = 40
            val_opts["headerHeight"] = 80
            val_height = compute_grid_height(filtered_val, 40, 80)
            val_res = AgGrid(filtered_val,
                             gridOptions=val_opts,
                             update_mode=GridUpdateMode.VALUE_CHANGED,
                             data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
                             key="val_grid",
                             height=val_height,
                             use_container_width=True)
            st.session_state.value_dist_df = pd.DataFrame(val_res["data"]).copy()
            
            # View SQL Logic for Value Distribution
            st.subheader("View SQL Logic (Value Distribution)")
            val_orig = st.session_state.df_data[st.session_state.df_data["analysis_type"]=="value_dist"]
            val_orig_field = val_orig[val_orig["field_name"]==selected_val_field]
            val_val_labels = val_orig_field["value_label"].dropna().unique().tolist()
            if val_val_labels:
                default_val_val = st.session_state.get("preselect_val_label_val", None)
                if default_val_val not in val_val_labels:
                    default_val_val = val_val_labels[0]
                sel_val_val_label = st.selectbox("Select Value Label (Value Dist)", val_val_labels, index=val_val_labels.index(default_val_val) if default_val_val else 0, key="val_sql_val_label")
                months_val = [(st.session_state.date1 - relativedelta(months=i)).replace(day=1) for i in range(12)]
                months_val = sorted(months_val, reverse=True)
                month_options_val = [m.strftime("%Y-%m") for m in months_val]
                sel_val_month = st.selectbox("Select Month (Value Dist)", month_options_val, key="val_sql_month")
                if st.button("Show SQL Logic (Value Dist)"):
                    matches_val = val_orig_field[
                        (val_orig_field["value_label"]==sel_val_val_label) &
                        (val_orig_field["filemonth_dt"].dt.strftime("%Y-%m")==sel_val_month)
                    ]
                    sql_vals_val = matches_val["value_sql_logic"].dropna().unique()
                    if sql_vals_val.size > 0:
                        st.text_area("Value SQL Logic (Value Dist)", "\n".join(sql_vals_val), height=150)
                    else:
                        st.text_area("Value SQL Logic (Value Dist)", "No SQL Logic found", height=150)
        
        # ---------------------------
        # Display Population Comparison Grid with "Prev Comments" Column
        st.subheader("Population Comparison")
        pop_fields = st.session_state.pop_comp_df["Field Name"].unique().tolist()
        if not pop_fields:
            st.warning("No Population Comparison data available.")
        else:
            active_pop = st.session_state.active_field if st.session_state.active_field in pop_fields else pop_fields[0]
            selected_pop_field = st.selectbox("Select Field (Pop Comp)", pop_fields, index=pop_fields.index(active_pop) if active_pop in pop_fields else 0, key="pop_field_select")
            st.session_state.active_field = selected_pop_field
            filtered_pop = st.session_state.pop_comp_df[st.session_state.pop_comp_df["Field Name"] == selected_pop_field].copy()
            if "Comment" not in filtered_pop.columns:
                filtered_pop["Comment"] = ""
            if "Prev Comments" in filtered_pop.columns:
                filtered_pop = filtered_pop.drop(columns=["Prev Comments"])
            pivot_prev = pivot_previous_comments(prev_comments_df, target_prev_month)
            if not pivot_prev.empty:
                filtered_pop = filtered_pop.merge(pivot_prev, on="Field Name", how="left")
                filtered_pop.rename(columns={f"comment_{target_prev_month}": "Prev Comments"}, inplace=True)
            else:
                filtered_pop["Prev Comments"] = ""
            gb_pop = GridOptionsBuilder.from_dataframe(filtered_pop)
            gb_pop.configure_default_column(editable=True, cellStyle={'white-space': 'normal', 'line-height': '1.2em', 'width': 150})
            gb_pop.configure_column("Comment", editable=True, width=150, minWidth=100, maxWidth=200)
            gb_pop.configure_column("Prev Comments", editable=False, width=150, minWidth=100, maxWidth=200)
            gb_pop.configure_selection("single", use_checkbox=True, suppressRowClickSelection=True)
            pop_opts = gb_pop.build()
            if isinstance(pop_opts, list):
                pop_opts = {"columnDefs": pop_opts}
            pop_opts["rowSelection"] = "single"
            pop_opts["pagination"] = False
            pop_opts["rowHeight"] = 40
            pop_opts["headerHeight"] = 80
            pop_height = compute_grid_height(filtered_pop, 40, 80)
            pop_res = AgGrid(filtered_pop,
                             gridOptions=pop_opts,
                             update_mode=GridUpdateMode.VALUE_CHANGED,
                             data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
                             key="pop_grid",
                             height=pop_height,
                             use_container_width=True)
            st.session_state.pop_comp_df = pd.DataFrame(pop_res["data"]).copy()
            
            # View SQL Logic for Population Comparison
            st.subheader("View SQL Logic (Population Comparison)")
            pop_orig = st.session_state.df_data[st.session_state.df_data["analysis_type"]=="pop_comp"]
            pop_orig_field = pop_orig[pop_orig["field_name"]==selected_pop_field]
            pop_val_labels = pop_orig_field["value_label"].dropna().unique().tolist()
            if pop_val_labels:
                default_pop_val = st.session_state.get("preselect_val_label_pop", None)
                if default_pop_val not in pop_val_labels:
                    default_pop_val = pop_val_labels[0]
                sel_pop_val_label = st.selectbox("Select Value Label (Pop Comp)", pop_val_labels, index=pop_val_labels.index(default_pop_val) if default_pop_val else 0, key="pop_sql_val_label")
                months = [(st.session_state.date1 - relativedelta(months=i)).replace(day=1) for i in range(12)]
                months = sorted(months, reverse=True)
                month_options = [m.strftime("%Y-%m") for m in months]
                sel_pop_month = st.selectbox("Select Month (Pop Comp)", month_options, key="pop_sql_month")
                if st.button("Show SQL Logic (Pop Comp)"):
                    matches = pop_orig_field[
                        (pop_orig_field["value_label"]==sel_pop_val_label) &
                        (pop_orig_field["filemonth_dt"].dt.strftime("%Y-%m")==sel_pop_month)
                    ]
                    sql_vals = matches["value_sql_logic"].dropna().unique()
                    if sql_vals.size > 0:
                        st.text_area("Value SQL Logic (Pop Comp)", "\n".join(sql_vals), height=150)
                    else:
                        st.text_area("Value SQL Logic (Pop Comp)", "No SQL Logic found", height=150)
        
        # ---------------------------
        # Preserve existing Summary comments from the input Excel file.
        summary_df = st.session_state.summary_df
        try:
            wb = load_workbook(st.session_state.input_file_path, data_only=True)
            if "Summary" in wb.sheetnames:
                existing = pd.read_excel(st.session_state.input_file_path, sheet_name="Summary")
                if "Approval Comments" in existing.columns:
                    summary_df = summary_df.merge(existing[["Field Name", "Approval Comments"]], on="Field Name", how="left", suffixes=("", "_old"))
                    summary_df["Approval Comments"] = summary_df.apply(lambda r: r["Approval Comments_old"] if pd.notnull(r["Approval Comments_old"]) else r["Approval Comments"], axis=1)
                    summary_df.drop(columns=["Approval Comments_old"], inplace=True)
                if "Comment" in existing.columns:
                    summary_df = summary_df.merge(existing[["Field Name", "Comment"]], on="Field Name", how="left", suffixes=("", "_old"))
                    summary_df["Comment"] = summary_df.apply(lambda r: r["Comment_old"] if pd.notnull(r["Comment_old"]) and r["Comment_old"] != "" else r["Comment"], axis=1)
                    summary_df.drop(columns=["Comment_old"], inplace=True)
        except Exception as e:
            st.warning(f"Could not preserve existing summary comments: {e}")
        st.session_state.summary_df = summary_df
        
        # ---------------------------
        # Aggregate current grid comments into Summary (Approval Comments remain unchanged)
        def aggregate_current_comments():
            sum_df = st.session_state.summary_df.copy()
            for field in sum_df["Field Name"].unique():
                notes = []
                if "Value Label" in st.session_state.value_dist_df.columns:
                    dist_df = st.session_state.value_dist_df[st.session_state.value_dist_df["Field Name"] == field]
                    for _, row in dist_df.iterrows():
                        comment = str(row.get("Comment", "")).strip()
                        val_label = str(row.get("Value Label", "")).strip()
                        if comment:
                            notes.append(f"{val_label} - {comment}")
                if "Value Label" in st.session_state.pop_comp_df.columns:
                    pop_df = st.session_state.pop_comp_df[st.session_state.pop_comp_df["Field Name"] == field]
                    for _, row in pop_df.iterrows():
                        comment = str(row.get("Comment", "")).strip()
                        val_label = str(row.get("Value Label", "")).strip()
                        if comment:
                            notes.append(f"{val_label} - {comment}")
                aggregated_note = "\n".join(notes)
                sum_df.loc[sum_df["Field Name"] == field, "Comment"] = aggregated_note
            st.session_state.summary_df = sum_df
        aggregate_current_comments()
        
        # ---------------------------
        # Reorder Summary columns so that Approval Comments comes immediately after Comment.
        sum_df = st.session_state.summary_df.copy()
        cols = list(sum_df.columns)
        cols.remove("Approval Comments")
        cols.remove("Comment")
        new_order = ["Field Name"] + cols[1:] + ["Comment", "Approval Comments"]
        sum_df = sum_df[new_order]
        
        # Display updated Summary Grid.
        st.subheader("Summary")
        gb_sum = GridOptionsBuilder.from_dataframe(sum_df)
        gb_sum.configure_default_column(editable=False,
                                        cellStyle={'white-space': 'normal', 'line-height': '1.2em', 'width': 150})
        gb_sum.configure_column("Approval Comments", editable=True, width=250, minWidth=100, maxWidth=300)
        gb_sum.configure_column("Comment", editable=False, width=250, minWidth=100, maxWidth=300)
        for col in sum_df.columns:
            if col not in ["Field Name", "Comment", "Approval Comments"]:
                if "Change" in col:
                    gb_sum.configure_column(
                        col,
                        type=["numericColumn"],
                        valueFormatter="(params.value != null ? params.value.toFixed(2)+'%' : '')",
                        width=150, minWidth=100, maxWidth=200
                    )
                else:
                    gb_sum.configure_column(
                        col,
                        type=["numericColumn"],
                        valueFormatter="(params.value != null ? params.value.toLocaleString('en-US') : '')",
                        width=150, minWidth=100, maxWidth=200
                    )
        sum_opts = gb_sum.build()
        if isinstance(sum_opts, list):
            sum_opts = {"columnDefs": sum_opts}
        for c in sum_opts["columnDefs"]:
            c["headerName"] = "\n".join(c["headerName"].split())
        sum_opts["rowSelection"] = "single"
        sum_opts["pagination"] = False
        sum_opts["rowHeight"] = 40
        sum_opts["headerHeight"] = 80
        sum_height = compute_grid_height(sum_df, 40, 80)
        AgGrid(sum_df,
               gridOptions=sum_opts,
               update_mode=GridUpdateMode.VALUE_CHANGED,
               data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
               key="sum_grid",
               height=sum_height,
               use_container_width=True)
        
        # ---------------------------
        # In-Place Update of the Input Excel File with Timestamp Message
        try:
            with pd.ExcelWriter(st.session_state.input_file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                # --- Summary Sheet ---
                export_sum = st.session_state.summary_df.copy().reset_index(drop=True)
                export_sum.to_excel(writer, index=False, sheet_name="Summary")
                summary_sheet = writer.sheets["Summary"]
                d1_col_name = f"Month-to-Month Diff ({st.session_state.date1.strftime('%Y-%m-%d')})"
                sum_cols = export_sum.columns.tolist()
                try:
                    d1_col_index = sum_cols.index(d1_col_name) + 1
                except ValueError:
                    d1_col_index = export_sum.shape[1]
                pivot_prev = pivot_previous_comments(get_cached_previous_comments(st.session_state.folder), target_prev_month)
                for idx, row in export_sum.iterrows():
                    field = row["Field Name"]
                    if target_prev_month in pivot_prev.columns:
                        prev_comment = pivot_prev.loc[pivot_prev["Field Name"]==field, f"comment_{target_prev_month}"]
                        if not prev_comment.empty:
                            prev_comment_str = str(prev_comment.values[0]).strip()
                            if prev_comment_str:
                                cell = summary_sheet.cell(row=idx+2, column=d1_col_index)
                                com_obj = Comment(prev_comment_str, "Prev")
                                com_obj.visible = True
                                cell.comment = com_obj
                
                # --- Value Distribution Sheet ---
                vd_df = st.session_state.value_dist_df.copy().reset_index(drop=True)
                vd_df.to_excel(writer, index=False, sheet_name="Value Distribution")
                vd_sheet = writer.sheets["Value Distribution"]
                vd_cols = vd_df.columns.tolist()
                date_str = st.session_state.date1.strftime("%Y-%m")
                try:
                    vd_sum_col_index = vd_cols.index(date_str + " Sum") + 1
                    vd_percent_col_index = vd_cols.index(date_str + " Percent") + 1
                except ValueError:
                    vd_sum_col_index = vd_percent_col_index = None
                for idx, row in vd_df.iterrows():
                    excel_row = idx + 2
                    field = row["Field Name"]
                    if target_prev_month in pivot_prev.columns:
                        prev_comment = pivot_prev.loc[pivot_prev["Field Name"]==field, f"comment_{target_prev_month}"]
                        if not prev_comment.empty:
                            prev_comment_str = str(prev_comment.values[0]).strip()
                            if prev_comment_str:
                                if vd_sum_col_index is not None:
                                    cell = vd_sheet.cell(row=excel_row, column=vd_sum_col_index)
                                    com_obj = Comment(prev_comment_str, "Prev")
                                    com_obj.visible = True
                                    cell.comment = com_obj
                                if vd_percent_col_index is not None:
                                    cell = vd_sheet.cell(row=excel_row, column=vd_percent_col_index)
                                    com_obj = Comment(prev_comment_str, "Prev")
                                    com_obj.visible = True
                                    cell.comment = com_obj
                
                # --- Population Comparison Sheet ---
                pop_df = st.session_state.pop_comp_df.copy().reset_index(drop=True)
                pop_df.to_excel(writer, index=False, sheet_name="Population Comparison")
                pop_sheet = writer.sheets["Population Comparison"]
                pop_cols = pop_df.columns.tolist()
                try:
                    pop_sum_col_index = pop_cols.index(date_str + " Sum") + 1
                    pop_percent_col_index = pop_cols.index(date_str + " Percent") + 1
                except ValueError:
                    pop_sum_col_index = pop_percent_col_index = None
                for idx, row in pop_df.iterrows():
                    excel_row = idx + 2
                    field = row["Field Name"]
                    if target_prev_month in pivot_prev.columns:
                        prev_comment = pivot_prev.loc[pivot_prev["Field Name"]==field, f"comment_{target_prev_month}"]
                        if not prev_comment.empty:
                            prev_comment_str = str(prev_comment.values[0]).strip()
                            if prev_comment_str:
                                if pop_sum_col_index is not None:
                                    cell = pop_sheet.cell(row=excel_row, column=pop_sum_col_index)
                                    com_obj = Comment(prev_comment_str, "Prev")
                                    com_obj.visible = True
                                    cell.comment = com_obj
                                if pop_percent_col_index is not None:
                                    cell = pop_sheet.cell(row=excel_row, column=pop_percent_col_index)
                                    com_obj = Comment(prev_comment_str, "Prev")
                                    com_obj.visible = True
                                    cell.comment = com_obj
            update_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            st.success(f"The input Excel file was last updated at {update_time}.")
        except Exception as e:
            st.error(f"Error updating the Excel file: {e}")

if __name__ == "__main__":
    main()