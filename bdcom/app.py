import streamlit as st
st.set_page_config(page_title="Final FRY14M Field Analysis", layout="wide", initial_sidebar_state="expanded")

import pandas as pd
import os
import datetime
import re
from io import BytesIO
from dateutil.relativedelta import relativedelta
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode
from st_aggrid.shared import JsCode
from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Font

#############################################
# Helper Functions
#############################################

def compute_grid_height(df, row_height=40, header_height=80):
    n = len(df)
    return header_height + (min(n, 30) * row_height)

def get_excel_engine(file_path):
    return 'pyxlsb' if file_path.lower().endswith('.xlsb') else None

def normalize_columns(df, mapping={"field_name": "Field Name", "value_label": "Value Label"}):
    # Strip whitespace from column names and rename according to mapping.
    df.columns = [str(col).strip() for col in df.columns]
    for orig, new in mapping.items():
        for col in df.columns:
            if col.lower() == orig.lower() and col != new:
                df.rename(columns={col: new}, inplace=True)
    return df

def flatten_dataframe(df):
    if isinstance(df.columns, pd.MultiIndex):
        df = df.reset_index()
        df.columns = [" ".join(map(str, col)).strip() if isinstance(col, tuple) else str(col).strip()
                      for col in df.columns.values]
    df = normalize_columns(df)
    return df

#############################################
# Generate Summary Sheet from Data
#############################################

def generate_summary_df(df_data, date1, date2):
    # Use original lowercase columns.
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
    # Create summary DataFrame with normalized column names.
    summary = pd.DataFrame(rows, columns=[
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
    summary["Missing % Change"] = summary.apply(lambda r: ((r[m1]-r[m2]) / r[m2] * 100) if r[m2]!=0 else None, axis=1)
    summary["Month-to-Month % Change"] = summary.apply(lambda r: ((r[d1]-r[d2]) / r[d2] * 100) if r[d2]!=0 else None, axis=1)
    new_order = ["Field Name", m1, m2, "Missing % Change", d1, d2, "Month-to-Month % Change"]
    summary = summary[new_order]
    summary["Comment"] = ""
    summary["Approval Comments"] = ""
    return summary

#############################################
# Generate Distribution/Population Comparison with Monthly Comment Columns
#############################################

def generate_distribution_df(df, analysis_type, date1):
    # This function uses original lowercase column names.
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

#############################################
# Pivot All Previous Comments by Month
#############################################

def pivot_all_previous_comments(df):
    # Pivot such that each unique Month becomes a column "YYYY-MM m- Comments"
    if df.empty:
        return pd.DataFrame()
    months = sorted(df["Month"].unique())
    result = None
    for month in months:
        grouped = df[df["Month"]==month].groupby("Field Name")["Comment"].apply(
            lambda x: "\n".join(x.dropna().astype(str).str.strip())
        ).reset_index()
        colname = f"{month} m- Comments"
        grouped = grouped.rename(columns={"Comment": colname})
        if result is None:
            result = grouped
        else:
            result = pd.merge(result, grouped, on="Field Name", how="outer")
    if result is None:
        result = pd.DataFrame()
    return result

#############################################
# Load or Generate Report Data
#############################################

def load_report_data(file_path, date1, date2):
    df_data = pd.read_excel(file_path, sheet_name="Data")
    df_data["filemonth_dt"] = pd.to_datetime(df_data["filemonth_dt"])
    # Do not normalize here because generation functions expect lowercase names.
    wb = load_workbook(file_path, data_only=True)
    if "Summary" in wb.sheetnames:
        summary_df = pd.read_excel(file_path, sheet_name="Summary")
    else:
        summary_df = generate_summary_df(df_data, date1, date2)
    if "Value Distribution" in wb.sheetnames:
        val_dist_df = pd.read_excel(file_path, sheet_name="Value Distribution")
    else:
        val_dist_df = generate_distribution_df(df_data, "value_dist", date1)
    if "Population Comparison" in wb.sheetnames:
        pop_comp_df = pd.read_excel(file_path, sheet_name="Population Comparison")
    else:
        pop_comp_df = generate_distribution_df(df_data, "pop_comp", date1)
    return df_data, summary_df, val_dist_df, pop_comp_df

#############################################
# Cache Previous Comments
#############################################

def cache_previous_comments(current_folder):
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
                field_cell = row[0]
                if field_cell.value:
                    field_name = str(field_cell.value).strip()
                    cell = row[col_index - 1]
                    comment_text = ""
                    if cell.comment and cell.comment.text is not None:
                        raw = str(cell.comment.text).strip()
                        if raw.lower() != "nan" and raw != "":
                            comment_text = raw
                    if comment_text:
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

#############################################
# Preserve Existing Summary Comments
#############################################

def preserve_summary_comments(input_file_path, summary_df):
    try:
        existing = pd.read_excel(input_file_path, sheet_name="Summary")
        comment_dict = existing.set_index("Field Name")["Comment"].to_dict() if "Comment" in existing.columns else {}
        approval_dict = existing.set_index("Field Name")["Approval Comments"].to_dict() if "Approval Comments" in existing.columns else {}
        summary_df["Comment"] = summary_df["Field Name"].map(comment_dict).fillna("")
        summary_df["Approval Comments"] = summary_df["Field Name"].map(approval_dict).fillna("")
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
    all_files = [f for f in os.listdir(folder_path) if f.lower().endswith(('.xlsx','.xlsb'))]
    if not all_files:
        st.sidebar.error(f"No Excel files found in folder '{folder}'.")
        return
    selected_file = st.sidebar.selectbox("Select an Excel File", all_files)
    input_file_path = os.path.join(folder_path, selected_file)
    selected_date = st.sidebar.date_input("Select Date for Date1", datetime.date(2025, 1, 1))
    date1 = datetime.datetime.combine(selected_date, datetime.datetime.min.time())
    date2 = date1 - relativedelta(months=1)
    
    # Cache previous comments and pivot them (all months in separate columns)
    prev_comments_df = get_cached_previous_comments(folder)
    if prev_comments_df.empty:
        prev_comments_df = cache_previous_comments(folder)
    st.write("Cached previous comments:")
    st.dataframe(prev_comments_df)
    pivot_prev_all = None
    # Pivot all previous comments so each unique Month becomes its own column.
    if not prev_comments_df.empty:
        # Helper: pivot all previous comments.
        def pivot_all_previous_comments(df):
            months = sorted(df["Month"].unique())
            result = None
            for month in months:
                grp = df[df["Month"]==month].groupby("Field Name")["Comment"].apply(
                    lambda x: "\n".join(x.dropna().astype(str).str.strip())
                ).reset_index()
                colname = f"{month} m- Comments"
                grp = grp.rename(columns={"Comment": colname})
                if result is None:
                    result = grp
                else:
                    result = pd.merge(result, grp, on="Field Name", how="outer")
            if result is None:
                result = pd.DataFrame()
            return result
        pivot_prev_all = pivot_all_previous_comments(prev_comments_df)
    
    if st.sidebar.button("Generate Report"):
        df_data, summary_df, val_dist_df, pop_comp_df = load_report_data(input_file_path, date1, date2)
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
        
        ##############################
        # Value Distribution Grid with previous comment columns
        ##############################
        st.subheader("Value Distribution")
        # Ensure normalized column name in the value distribution DataFrame.
        if "Field Name" not in st.session_state.value_dist_df.columns:
            st.session_state.value_dist_df = normalize_columns(st.session_state.value_dist_df)
        try:
            val_fields = st.session_state.value_dist_df["Field Name"].unique().tolist()
        except KeyError:
            val_fields = st.session_state.df_data["field_name"].unique().tolist()
            st.warning("Column 'Field Name' not found in Value Distribution data; using Data sheet.")
        if not val_fields:
            st.warning("No Value Distribution data available.")
        else:
            active_val = st.session_state.active_field if st.session_state.active_field in val_fields else val_fields[0]
            selected_val_field = st.selectbox("Select Field (Value Dist)", val_fields,
                                              index=val_fields.index(active_val),
                                              key="val_field_select")
            st.session_state.active_field = selected_val_field
            filtered_val = st.session_state.value_dist_df[st.session_state.value_dist_df["Field Name"] == selected_val_field].copy()
            # Merge pivoted previous comments if available.
            if pivot_prev_all is not None and not pivot_prev_all.empty:
                filtered_val = pd.merge(filtered_val, pivot_prev_all, on="Field Name", how="left")
            gb_val = GridOptionsBuilder.from_dataframe(filtered_val)
            gb_val.configure_default_column(editable=True, cellStyle={'white-space':'normal','line-height':'1.2em','width':150})
            # For the "Comment" column in the grid, use a custom cellRenderer to render HTML.
            custom_renderer = JsCode("function(params){return params.value ? params.value : '';}")
            gb_val.configure_column("Comment", editable=True, cellRenderer=custom_renderer, width=180)
            for c in filtered_val.columns:
                if "m- Comments" in c:
                    gb_val.configure_column(c, editable=False, width=150)
                elif c.endswith("Sum") or c.endswith("Percent"):
                    gb_val.configure_column(c, editable=False, width=120)
            val_opts = gb_val.build()
            if isinstance(val_opts, list):
                val_opts = {"columnDefs": val_opts}
            val_opts["rowSelection"] = "single"
            val_opts["pagination"] = False
            val_opts["rowHeight"] = 40
            val_opts["headerHeight"] = 80
            val_height = compute_grid_height(filtered_val, 40, 80)
            val_res = AgGrid(filtered_val, gridOptions=val_opts,
                             update_mode=GridUpdateMode.VALUE_CHANGED,
                             data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
                             key="val_grid", height=val_height, use_container_width=True)
            st.session_state.value_dist_df = pd.DataFrame(val_res["data"]).copy()
            
            # View SQL Logic for Value Distribution (using original lowercase columns)
            st.subheader("View SQL Logic (Value Distribution)")
            val_orig = st.session_state.df_data[st.session_state.df_data["analysis_type"]=="value_dist"]
            val_orig_field = val_orig[val_orig["field_name"]==selected_val_field]
            val_val_labels = val_orig_field["value_label"].dropna().unique().tolist()
            if val_val_labels:
                default_val_val = st.session_state.get("preselect_val_label_val", None)
                if default_val_val not in val_val_labels:
                    default_val_val = val_val_labels[0]
                sel_val_val_label = st.selectbox("Select Value Label (Value Dist)", val_val_labels,
                                                 index=val_val_labels.index(default_val_val) if default_val_val else 0,
                                                 key="val_sql_val_label")
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
        
        ##############################
        # Population Comparison Grid with previous comment columns
        ##############################
        st.subheader("Population Comparison")
        if "Field Name" not in st.session_state.pop_comp_df.columns:
            st.session_state.pop_comp_df = normalize_columns(st.session_state.pop_comp_df)
        try:
            pop_fields = st.session_state.pop_comp_df["Field Name"].unique().tolist()
        except KeyError:
            pop_fields = st.session_state.df_data["field_name"].unique().tolist()
            st.warning("Column 'Field Name' not found in Population Comparison data; using Data sheet.")
        if not pop_fields:
            st.warning("No Population Comparison data available.")
        else:
            active_pop = st.session_state.active_field if st.session_state.active_field in pop_fields else pop_fields[0]
            selected_pop_field = st.selectbox("Select Field (Pop Comp)", pop_fields,
                                              index=pop_fields.index(active_pop) if active_pop in pop_fields else 0,
                                              key="pop_field_select")
            st.session_state.active_field = selected_pop_field
            filtered_pop = st.session_state.pop_comp_df[st.session_state.pop_comp_df["Field Name"] == selected_pop_field].copy()
            if "Prev Comments" in filtered_pop.columns:
                filtered_pop = filtered_pop.drop(columns=["Prev Comments"])
            if pivot_prev_all is not None and not pivot_prev_all.empty:
                filtered_pop = pd.merge(filtered_pop, pivot_prev_all, on="Field Name", how="left")
                # Rename the pivoted column(s) if needed.
            gb_pop = GridOptionsBuilder.from_dataframe(filtered_pop)
            gb_pop.configure_default_column(editable=True, cellStyle={'white-space':'normal','line-height':'1.2em','width':150})
            gb_pop.configure_column("Comment", editable=True, width=180)
            for c in filtered_pop.columns:
                if "m- Comments" in c:
                    gb_pop.configure_column(c, editable=False, width=150)
                elif c.endswith("Sum") or c.endswith("Percent"):
                    gb_pop.configure_column(c, editable=False, width=120)
            pop_opts = gb_pop.build()
            if isinstance(pop_opts, list):
                pop_opts = {"columnDefs": pop_opts}
            pop_opts["rowSelection"] = "single"
            pop_opts["pagination"] = False
            pop_opts["rowHeight"] = 40
            pop_opts["headerHeight"] = 80
            pop_height = compute_grid_height(filtered_pop, 40, 80)
            pop_res = AgGrid(filtered_pop, gridOptions=pop_opts,
                             update_mode=GridUpdateMode.VALUE_CHANGED,
                             data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
                             key="pop_grid", height=pop_height, use_container_width=True)
            st.session_state.pop_comp_df = pd.DataFrame(pop_res["data"]).copy()
            
            st.subheader("View SQL Logic (Population Comparison)")
            pop_orig = st.session_state.df_data[st.session_state.df_data["analysis_type"]=="pop_comp"]
            pop_orig_field = pop_orig[pop_orig["field_name"]==selected_pop_field]
            pop_val_labels = pop_orig_field["value_label"].dropna().unique().tolist()
            if pop_val_labels:
                default_pop_val = st.session_state.get("preselect_val_label_pop", None)
                if default_pop_val not in pop_val_labels:
                    default_pop_val = pop_val_labels[0]
                sel_pop_val_label = st.selectbox("Select Value Label (Pop Comp)", pop_val_labels,
                                                 index=pop_val_labels.index(default_pop_val) if default_pop_val else 0,
                                                 key="pop_sql_val_label")
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
        
        ##############################
        # Append Previous Month Comment in red to Summary "Comment" column
        ##############################
        # Pivot previous comments for target_prev_month:
        pivot_prev = pivot_previous_comments(get_cached_previous_comments(st.session_state.folder), target_prev_month)
        # For each row in summary, if there's a previous comment, append it in HTML with red color.
        def append_prev(existing, field):
            if pivot_prev.empty:
                return existing
            prev_val = pivot_prev[pivot_prev["Field Name"]==field]
            if not prev_val.empty:
                # Get the previous comment text from the appropriate column (e.g., "comment_2024-12")
                colname = list(prev_val.columns[1:])[0]  # assume only one column is present
                prev_comment = prev_val.iloc[0][colname]
                if prev_comment and str(prev_comment).strip() != "":
                    if pd.isna(existing) or existing=="":
                        return f"<span style='color:red;'>{prev_comment}</span>"
                    else:
                        return existing + "\n<span style='color:red;'>{}</span>".format(prev_comment)
            return existing
        
        summary_df = st.session_state.summary_df.copy()
        summary_df["Comment"] = summary_df.apply(lambda r: append_prev(r["Comment"], r["Field Name"]), axis=1)
        st.session_state.summary_df = summary_df
        
        ##############################
        # Preserve existing Summary comments from the input Excel file.
        try:
            existing = pd.read_excel(st.session_state.input_file_path, sheet_name="Summary")
            comment_dict = existing.set_index("Field Name")["Comment"].to_dict() if "Comment" in existing.columns else {}
            approval_dict = existing.set_index("Field Name")["Approval Comments"].to_dict() if "Approval Comments" in existing.columns else {}
            st.session_state.summary_df["Comment"] = st.session_state.summary_df["Field Name"].map(comment_dict).fillna("")\
                .str.cat(st.session_state.summary_df["Comment"], sep="\n")
            st.session_state.summary_df["Approval Comments"] = st.session_state.summary_df["Field Name"].map(approval_dict).fillna("")
        except Exception as e:
            st.warning(f"Could not preserve existing summary comments: {e}")
        
        ##############################
        # Aggregate current grid comments into Summary's "Comment" Column
        ##############################
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
                if aggregated_note:
                    sum_df.loc[sum_df["Field Name"] == field, "Comment"] = aggregated_note
            st.session_state.summary_df = sum_df
        aggregate_current_comments()
        
        # Ensure "Approval Comments" and "Comment" columns exist
        if "Approval Comments" not in st.session_state.summary_df.columns:
            st.session_state.summary_df["Approval Comments"] = ""
        if "Comment" not in st.session_state.summary_df.columns:
            st.session_state.summary_df["Comment"] = ""
        
        # Reorder Summary columns so that "Approval Comments" comes immediately after "Comment"
        sum_df = st.session_state.summary_df.copy()
        cols = list(sum_df.columns)
        if "Approval Comments" in cols:
            cols.remove("Approval Comments")
        if "Comment" in cols:
            cols.remove("Comment")
        new_order = ["Field Name"] + [c for c in cols if c != "Field Name"] + ["Comment", "Approval Comments"]
        sum_df = sum_df[new_order]
        
        st.subheader("Summary")
        # Use a custom cellRenderer for the "Comment" column in AgGrid so that HTML is rendered.
        comment_renderer = JsCode("function(params){return params.value ? params.value : '';}")
        gb_sum = GridOptionsBuilder.from_dataframe(sum_df)
        gb_sum.configure_default_column(editable=False, cellStyle={'white-space':'normal','line-height':'1.2em','width':150})
        gb_sum.configure_column("Approval Comments", editable=True, width=250, minWidth=100, maxWidth=300)
        gb_sum.configure_column("Comment", editable=False, cellRenderer=comment_renderer, width=250, minWidth=100, maxWidth=300)
        for c in sum_df.columns:
            if c not in ["Field Name", "Comment", "Approval Comments"]:
                if "Change" in c:
                    gb_sum.configure_column(c, type=["numericColumn"],
                                            valueFormatter="(params.value != null ? params.value.toFixed(2)+'%' : '')",
                                            width=150, minWidth=100, maxWidth=200)
                else:
                    gb_sum.configure_column(c, type=["numericColumn"],
                                            valueFormatter="(params.value != null ? params.value.toLocaleString('en-US') : '')",
                                            width=150, minWidth=100, maxWidth=200)
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
        sum_res = AgGrid(sum_df, gridOptions=sum_opts,
                         update_mode=GridUpdateMode.VALUE_CHANGED,
                         data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
                         key="sum_grid", height=sum_height, use_container_width=True)
        st.session_state.summary_df = pd.DataFrame(sum_res["data"]).copy()
        
        ##############################
        # In-Place Update of the Input Excel File
        ##############################
        try:
            with pd.ExcelWriter(st.session_state.input_file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                st.session_state.summary_df.to_excel(writer, index=False, sheet_name="Summary")
                st.session_state.value_dist_df.to_excel(writer, index=False, sheet_name="Value Distribution")
                st.session_state.pop_comp_df.to_excel(writer, index=False, sheet_name="Population Comparison")
                # Now update the Comment cells in the Summary sheet to append previous month comment in red.
                summary_sheet = writer.sheets["Summary"]
                d1_col_name = f"Month-to-Month Diff ({st.session_state.date1.strftime('%Y-%m-%d')})"
                sum_cols = st.session_state.summary_df.columns.tolist()
                try:
                    d1_col_index = sum_cols.index(d1_col_name) + 1
                except ValueError:
                    d1_col_index = st.session_state.summary_df.shape[1]
                # For each row, if a previous month's comment exists (from pivot_prev_all for target_prev_month),
                # append it (in red) to the cell comment.
                pivot_prev = pivot_previous_comments(get_cached_previous_comments(st.session_state.folder), target_prev_month)
                for idx, row in st.session_state.summary_df.iterrows():
                    field = row["Field Name"]
                    if not pivot_prev.empty and field in pivot_prev["Field Name"].values:
                        prev_comment = pivot_prev[pivot_prev["Field Name"]==field].iloc[0,1]
                        if prev_comment and str(prev_comment).strip() != "":
                            # Construct new comment text: original comment + newline + previous comment in red.
                            orig_comment = row["Comment"] if row["Comment"] else ""
                            new_comment = orig_comment + "\n" + prev_comment if orig_comment else prev_comment
                            # Create a comment and set its font to red.
                            com_obj = Comment(new_comment, "User")
                            com_obj.font = Font(color="FF0000")
                            summary_sheet.cell(row=idx+2, column=d1_col_index).comment = com_obj
            update_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            st.success(f"Excel file updated successfully at {update_time}.")
        except Exception as e:
            st.error(f"Error updating Excel file: {e}")

if __name__ == "__main__":
    main()