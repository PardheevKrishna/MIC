import streamlit as st
st.set_page_config(page_title="FRY14M Field Analysis", layout="wide", initial_sidebar_state="expanded")

import pandas as pd
import os
import datetime
import re
from io import BytesIO
from dateutil.relativedelta import relativedelta
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode
from openpyxl import load_workbook
from openpyxl.comments import Comment
import numpy as np

#############################################
# Helper Functions
#############################################

def compute_grid_height(df, row_height=40, header_height=80):
    n = len(df)
    return header_height + (min(n, 30) * row_height)

def get_excel_engine(file_path):
    return 'pyxlsb' if file_path.lower().endswith('.xlsb') else None

def normalize_columns(df, mapping={"field_name": "Field Name", "value_label": "Value Label"}):
    # Strip whitespace and rename columns based on mapping.
    df.columns = [str(col).strip() for col in df.columns]
    for orig, new in mapping.items():
        for col in df.columns:
            if col.lower() == orig.lower() and col != new:
                df.rename(columns={col: new}, inplace=True)
    return df

def flatten_dataframe(df):
    """Flatten MultiIndex columns and ensure a column named 'Field Name' exists."""
    if isinstance(df.columns, pd.MultiIndex):
        df = df.reset_index()
        df.columns = [" ".join(map(str, col)).strip() if isinstance(col, tuple) else str(col).strip() 
                      for col in df.columns.values]
    # Ensure normalized columns are present.
    df = normalize_columns(df)
    # If "Field Name" is still missing, try to rename the first column if it appears unnamed.
    if "Field Name" not in df.columns:
        first_col = df.columns[0]
        if isinstance(first_col, tuple):
            first_col_str = " ".join(map(str, first_col)).strip()
        else:
            first_col_str = str(first_col).strip()
        if first_col_str.lower().startswith("unnamed") or first_col_str == "":
            df.rename(columns={df.columns[0]: "Field Name"}, inplace=True)
    return df

#############################################
# Generate Summary Sheet from Data
#############################################

def generate_summary_df(df_data, date1, date2):
    df_data = normalize_columns(df_data)
    fields = sorted(df_data["Field Name"].unique())
    rows = []
    for field in fields:
        mask_miss_d1 = ((df_data['analysis_type'] == 'value_dist') &
                        (df_data['Field Name'] == field) &
                        (df_data['filemonth_dt'] == date1) &
                        (df_data['Value Label'].str.contains("Missing", case=False, na=False)))
        missing_d1 = df_data.loc[mask_miss_d1, 'value_records'].sum()
        mask_miss_d2 = ((df_data['analysis_type'] == 'value_dist') &
                        (df_data['Field Name'] == field) &
                        (df_data['filemonth_dt'] == date2) &
                        (df_data['Value Label'].str.contains("Missing", case=False, na=False)))
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
                       (df_data['Field Name'] == field) &
                       (df_data['filemonth_dt'] == date1) &
                       (df_data['Value Label'].apply(contains_phrase)))
        pop_d1 = df_data.loc[mask_pop_d1, 'value_records'].sum()
        mask_pop_d2 = ((df_data['analysis_type'] == 'pop_comp') &
                       (df_data['Field Name'] == field) &
                       (df_data['filemonth_dt'] == date2) &
                       (df_data['Value Label'].apply(contains_phrase)))
        pop_d2 = df_data.loc[mask_pop_d2, 'value_records'].sum()
        rows.append([field, missing_d1, missing_d2, pop_d1, pop_d2])
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

def generate_dist_with_comments(df, analysis_type, date1):
    df = normalize_columns(df)
    months = [(date1 - relativedelta(months=i)).replace(day=1) for i in range(12)]
    months = sorted(months, reverse=True)
    sub = df[df['analysis_type'] == analysis_type].copy()
    sub['month'] = sub['filemonth_dt'].apply(lambda d: d.replace(day=1))
    sub = sub[sub['month'].isin(months)]
    if 'value_comment' not in sub.columns:
        sub['value_comment'] = ""
    grouped = sub.groupby(['Field Name', 'Value Label', 'month'])['value_records'].sum().reset_index()
    if grouped.empty:
        return pd.DataFrame()
    pivot_num = grouped.pivot_table(
        index=['Field Name', 'Value Label'],
        columns='month',
        values='value_records',
        fill_value=0,
        aggfunc='sum'
    )
    pivot_num = pivot_num.reindex(columns=months, fill_value=0)
    def first_non_empty(vals):
        for v in vals:
            if pd.notna(v) and str(v).strip().lower() != "nan" and str(v).strip() != "":
                return v
        return ""
    grouped_cmt = sub.groupby(['Field Name', 'Value Label', 'month'])['value_comment'].apply(first_non_empty).reset_index()
    pivot_cmt = grouped_cmt.pivot_table(
        index=['Field Name', 'Value Label'],
        columns='month',
        values='value_comment',
        fill_value="",
        aggfunc='first'
    )
    pivot_cmt = pivot_cmt.reindex(columns=months, fill_value="")
    frames = []
    for field, num_subdf in pivot_num.groupby(level=0):
        num_subdf = num_subdf.droplevel(0)
        if field in pivot_cmt.index.get_level_values(0):
            cmt_subdf = pivot_cmt.loc[field]
        else:
            cmt_subdf = pd.DataFrame()
        if isinstance(cmt_subdf, pd.Series):
            cmt_subdf = cmt_subdf.to_frame().T
        total = num_subdf.sum(axis=0)
        pct = num_subdf.div(total, axis=1).mul(100).round(2).fillna(0)
        combined = pd.DataFrame(index=num_subdf.index)
        for m in months:
            m_str = m.strftime('%Y-%m')
            combined[(m_str, "Sum")] = num_subdf[m]
            combined[(m_str, "Percent")] = pct[m]
            def get_comment(vlabel):
                if vlabel in cmt_subdf.index and m in cmt_subdf.columns:
                    val = cmt_subdf.loc[vlabel, m]
                    if pd.isna(val) or str(val).strip().lower() == "nan":
                        return ""
                    return str(val)
                return ""
            combined[(m_str, "Comment")] = [get_comment(vlabel) for vlabel in combined.index]
        tot_row = {}
        for m in months:
            m_str = m.strftime('%Y-%m')
            tot_row[(m_str, "Sum")] = total[m]
            tot_row[(m_str, "Percent")] = ""
            tot_row[(m_str, "Comment")] = ""
        combined.loc["Current period total"] = tot_row
        combined.index = pd.MultiIndex.from_product([[field], combined.index], names=["Field Name", "Value Label"])
        frames.append(combined)
    if not frames:
        return pd.DataFrame()
    final = pd.concat(frames)
    return final

#############################################
# Load or Generate Report Data
#############################################

def load_report_data(file_path, date1, date2):
    df_data = pd.read_excel(file_path, sheet_name="Data")
    df_data["filemonth_dt"] = pd.to_datetime(df_data["filemonth_dt"])
    df_data = normalize_columns(df_data)  # Ensure normalization right away.
    wb = load_workbook(file_path, data_only=True)
    if "Summary" in wb.sheetnames:
        summary_df = pd.read_excel(file_path, sheet_name="Summary")
        summary_df = normalize_columns(summary_df)
    else:
        summary_df = generate_summary_df(df_data, date1, date2)
    if "Value Distribution" in wb.sheetnames:
        val_dist_df = pd.read_excel(file_path, sheet_name="Value Distribution")
        val_dist_df = normalize_columns(val_dist_df)
    else:
        val_dist_df = generate_dist_with_comments(df_data, "value_dist", date1)
    if "Population Comparison" in wb.sheetnames:
        pop_comp_df = pd.read_excel(file_path, sheet_name="Population Comparison")
        pop_comp_df = normalize_columns(pop_comp_df)
    else:
        pop_comp_df = generate_dist_with_comments(df_data, "pop_comp", date1)
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

def pivot_previous_comments(df, target_month):
    if df.empty:
        return pd.DataFrame()
    df_target = df[df["Month"] == target_month]
    if df_target.empty:
        return pd.DataFrame()
    grouped = df_target.groupby(["Field Name"])["Comment"].apply(
        lambda x: "\n".join(
            str(item) for item in x if pd.notnull(item) and str(item).strip().lower() != "nan" and str(item).strip() != ""
        )
    ).reset_index()
    colname = f"{target_month} m- Comments"
    grouped = grouped.rename(columns={"Comment": colname})
    return grouped

#############################################
# Preserve Existing Summary Comments
#############################################

def preserve_summary_comments(input_file_path, summary_df):
    try:
        wb = load_workbook(input_file_path, data_only=True)
        if "Summary" in wb.sheetnames:
            existing = pd.read_excel(input_file_path, sheet_name="Summary")
            existing = normalize_columns(existing)
            comment_dict = {}
            approval_dict = {}
            if "Field Name" in existing.columns and "Comment" in existing.columns:
                comment_dict = existing.set_index("Field Name")["Comment"].to_dict()
            if "Field Name" in existing.columns and "Approval Comments" in existing.columns:
                approval_dict = existing.set_index("Field Name")["Approval Comments"].to_dict()
            def safe_comment(field):
                val = comment_dict.get(field, "")
                if pd.isna(val) or str(val).strip().lower() == "nan":
                    return ""
                return str(val)
            def safe_approval(field):
                val = approval_dict.get(field, "")
                if pd.isna(val) or str(val).strip().lower() == "nan":
                    return ""
                return str(val)
            summary_df["Comment"] = summary_df["Field Name"].map(safe_comment)
            summary_df["Approval Comments"] = summary_df["Field Name"].map(safe_approval)
        if "Approval Comments" not in summary_df.columns:
            summary_df["Approval Comments"] = ""
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

    # Cache previous comments
    prev_comments_df = get_cached_previous_comments(folder)
    if prev_comments_df.empty:
        prev_comments_df = cache_previous_comments(folder)
    st.write("Cached previous comments (all months):")
    st.dataframe(prev_comments_df)

    target_prev_month = (date1 - relativedelta(months=1)).strftime("%Y-%m")

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
        # Value Distribution Grid (Monthly Columns)
        ##############################
        st.subheader("Value Distribution (Monthly Columns)")
        if "Field Name" not in st.session_state.value_dist_df.columns:
            st.session_state.value_dist_df = normalize_columns(st.session_state.value_dist_df)
        try:
            val_fields = st.session_state.value_dist_df["Field Name"].unique().tolist()
        except KeyError:
            val_fields = st.session_state.df_data["Field Name"].unique().tolist()
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
            gb_val = GridOptionsBuilder.from_dataframe(filtered_val)
            gb_val.configure_default_column(editable=True, cellStyle={'white-space':'normal','line-height':'1.2em','width':150})
            for c in filtered_val.columns:
                if c.endswith("Comment"):
                    gb_val.configure_column(c, editable=True, width=180)
                elif c.endswith("Sum") or c.endswith("Percent"):
                    gb_val.configure_column(c, editable=False, width=120)
            val_opts = gb_val.build()
            if isinstance(val_opts, list):
                val_opts = {"columnDefs": val_opts}
            val_res = AgGrid(filtered_val, gridOptions=val_opts,
                             update_mode=GridUpdateMode.VALUE_CHANGED,
                             data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
                             key="val_grid", height=compute_grid_height(filtered_val), use_container_width=True)
            st.session_state.value_dist_df = pd.DataFrame(val_res["data"]).copy()

        ##############################
        # Population Comparison Grid (Monthly Columns)
        ##############################
        st.subheader("Population Comparison (Monthly Columns)")
        if "Field Name" not in st.session_state.pop_comp_df.columns:
            st.session_state.pop_comp_df = normalize_columns(st.session_state.pop_comp_df)
        try:
            pop_fields = st.session_state.pop_comp_df["Field Name"].unique().tolist()
        except KeyError:
            pop_fields = st.session_state.df_data["Field Name"].unique().tolist()
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
            gb_pop = GridOptionsBuilder.from_dataframe(filtered_pop)
            gb_pop.configure_default_column(editable=True, cellStyle={'white-space':'normal','line-height':'1.2em','width':150})
            for c in filtered_pop.columns:
                if c.endswith("Comment"):
                    gb_pop.configure_column(c, editable=True, width=180)
                elif c.endswith("Sum") or c.endswith("Percent"):
                    gb_pop.configure_column(c, editable=False, width=120)
            pop_opts = gb_pop.build()
            if isinstance(pop_opts, list):
                pop_opts = {"columnDefs": pop_opts}
            pop_res = AgGrid(filtered_pop, gridOptions=pop_opts,
                             update_mode=GridUpdateMode.VALUE_CHANGED,
                             data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
                             key="pop_grid", height=compute_grid_height(filtered_pop), use_container_width=True)
            st.session_state.pop_comp_df = pd.DataFrame(pop_res["data"]).copy()

        ##############################
        # Aggregate Monthly Comments into Summary's "Comment" Column
        ##############################
        def aggregate_current_comments():
            sum_df = st.session_state.summary_df.copy()
            if "Field Name" not in sum_df.columns:
                st.warning("'Field Name' column missing in Summary; skipping comment aggregation.")
                st.session_state.summary_df = sum_df
                return
            for field in sum_df["Field Name"].unique():
                notes = []
                try:
                    vdist = st.session_state.value_dist_df[st.session_state.value_dist_df["Field Name"] == field]
                except KeyError:
                    vdist = pd.DataFrame()
                if not vdist.empty:
                    cmt_cols = [c for c in vdist.columns if c.endswith("Comment")]
                    for _, row in vdist.iterrows():
                        val_label = row.get("Value Label", "")
                        for c in cmt_cols:
                            raw_comment = str(row.get(c, "")).strip()
                            if raw_comment.lower() == "nan" or raw_comment == "":
                                continue
                            month_name = c.replace(" Comment", "").strip()
                            line = f"{val_label} ({month_name}) - {raw_comment}".strip(" -")
                            notes.append(line)
                try:
                    pcomp = st.session_state.pop_comp_df[st.session_state.pop_comp_df["Field Name"] == field]
                except KeyError:
                    pcomp = pd.DataFrame()
                if not pcomp.empty:
                    cmt_cols2 = [c for c in pcomp.columns if c.endswith("Comment")]
                    for _, row in pcomp.iterrows():
                        val_label = row.get("Value Label", "")
                        for c in cmt_cols2:
                            raw_comment = str(row.get(c, "")).strip()
                            if raw_comment.lower() == "nan" or raw_comment == "":
                                continue
                            month_name = c.replace(" Comment", "").strip()
                            line = f"{val_label} ({month_name}) - {raw_comment}".strip(" -")
                            notes.append(line)
                aggregated_note = "\n".join(notes).strip()
                if aggregated_note:
                    sum_df.loc[sum_df["Field Name"] == field, "Comment"] = aggregated_note
            st.session_state.summary_df = sum_df

        aggregate_current_comments()

        # Ensure that both "Approval Comments" and "Comment" columns exist
        if "Approval Comments" not in st.session_state.summary_df.columns:
            st.session_state.summary_df["Approval Comments"] = ""
        if "Comment" not in st.session_state.summary_df.columns:
            st.session_state.summary_df["Comment"] = ""

        # Reorder Summary columns so that "Approval Comments" comes after "Comment"
        sum_df = st.session_state.summary_df.copy()
        cols = list(sum_df.columns)
        if "Approval Comments" in cols:
            cols.remove("Approval Comments")
        if "Comment" in cols:
            cols.remove("Comment")
        new_order = ["Field Name"] + [c for c in cols if c != "Field Name"] + ["Comment", "Approval Comments"]
        sum_df = sum_df[new_order]

        st.subheader("Summary")
        gb_sum = GridOptionsBuilder.from_dataframe(sum_df)
        gb_sum.configure_default_column(editable=False,
            cellStyle={'white-space':'normal','line-height':'1.2em','width':150})
        gb_sum.configure_column("Approval Comments", editable=True, width=250, minWidth=100, maxWidth=300)
        gb_sum.configure_column("Comment", editable=False, width=250, minWidth=100, maxWidth=300)
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
        # Create a new Excel file for download
        ##############################
        out_buf = BytesIO()
        with pd.ExcelWriter(out_buf, engine='openpyxl') as writer:
            st.session_state.summary_df.to_excel(writer, index=False, sheet_name="Summary")
            st.session_state.value_dist_df.to_excel(writer, index=False, sheet_name="Value Distribution")
            st.session_state.pop_comp_df.to_excel(writer, index=False, sheet_name="Population Comparison")
        out_buf.seek(0)
        st.download_button("Download Report as Excel",
                           data=out_buf.getvalue(),
                           file_name="FRY14M_Field_Analysis_Report.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    main()