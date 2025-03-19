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

#############################################
# Generating Summary
#############################################

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

    # Add columns for aggregated current grid comments + editable Approval Comments
    df["Comment"] = ""
    df["Approval Comments"] = ""
    return df

#############################################
# Generating Dist/Pcomp with monthly comment columns
#############################################

def generate_dist_with_comments(df, analysis_type, date1):
    """
    Creates a pivot with columns for each of the last 12 months in descending order:
      <month> Sum, <month> Percent, <month> Comment
    We'll assume 'value_comment' is a column in the raw data for that row. If not found, we default to "".
    """
    # Last 12 months in descending order
    months = [(date1 - relativedelta(months=i)).replace(day=1) for i in range(12)]
    months = sorted(months, reverse=True)

    # Filter for the correct analysis_type
    sub = df[df['analysis_type'] == analysis_type].copy()
    sub['month'] = sub['filemonth_dt'].apply(lambda d: d.replace(day=1))
    sub = sub[sub['month'].isin(months)]

    # We'll store numeric data in one pivot, and comments in a separate pivot
    # if the 'value_comment' column doesn't exist, create it as ""
    if 'value_comment' not in sub.columns:
        sub['value_comment'] = ""

    # Create a numeric pivot
    grouped = sub.groupby(['field_name', 'value_label', 'month'])['value_records'].sum().reset_index()
    if grouped.empty:
        # Return an empty df if no data
        return pd.DataFrame()

    pivot_num = grouped.pivot_table(index=['field_name', 'value_label'],
                                    columns='month', values='value_records', fill_value=0)
    pivot_num = pivot_num.reindex(columns=months, fill_value=0)

    # Create a comment pivot
    # For each (field_name, value_label, month), store the 'value_comment' (the last row if multiple?)
    # We'll do a groupby and get the first or join them. We'll just pick the first non-empty.
    def first_non_empty(vals):
        for v in vals:
            if pd.notna(v) and str(v).strip().lower() != "nan" and v != "":
                return v
        return ""

    grouped_cmt = sub.groupby(['field_name', 'value_label', 'month'])['value_comment'].apply(first_non_empty).reset_index()
    pivot_cmt = grouped_cmt.pivot_table(index=['field_name', 'value_label'],
                                        columns='month', values='value_comment', fill_value="")
    pivot_cmt = pivot_cmt.reindex(columns=months, fill_value="")

    # Now we combine numeric pivot (Sum, Percent) with comment pivot
    frames = []
    for field, num_subdf in pivot_num.groupby(level=0):
        # numeric subdf for this field
        num_subdf = num_subdf.droplevel(0)  # remove field name from index
        # comment subdf for this field
        cmt_subdf = pivot_cmt.loc[field].to_frame().T if field in pivot_cmt.index else None
        # cmt_subdf might be a DataFrame with index= value_label, columns= months
        if cmt_subdf is None or cmt_subdf.empty:
            cmt_subdf = pivot_cmt.loc[[field]]  # empty?

        total = num_subdf.sum(axis=0)
        pct = num_subdf.div(total, axis=1).mul(100).round(2).fillna(0)

        data = {}
        for m in months:
            m_str = m.strftime('%Y-%m')
            # numeric sums
            data[(m_str, "Sum")] = num_subdf[m]
            # numeric percent
            data[(m_str, "Percent")] = pct[m]
            # comment
            # we need the row index to match each value_label
            # so let's do row by row below. We'll build a DataFrame row by row.
        # We'll build row by row after we build the top-level columns.

        # We'll create an empty DataFrame with the same row index as num_subdf
        combined = pd.DataFrame(index=num_subdf.index)

        # For each month, create 3 columns: (m_str, "Sum"), (m_str, "Percent"), (m_str, "Comment")
        # Then fill them from num_subdf, cmt_subdf
        for m in months:
            m_str = m.strftime('%Y-%m')
            combined[(m_str, "Sum")] = num_subdf[m]
            # compute percent
            combined[(m_str, "Percent")] = pct[m]
            # comment
            # For each row (value_label), get the comment from cmt_subdf
            # cmt_subdf is shaped [value_label x months]
            # so cmt_subdf.loc[value_label, m]
            def get_comment(vlabel):
                if vlabel in cmt_subdf.index:
                    val = cmt_subdf.loc[vlabel, m]
                    if pd.isna(val) or str(val).strip().lower() == "nan":
                        return ""
                    return str(val)
                return ""
            combined[(m_str, "Comment")] = [get_comment(vlabel) for vlabel in combined.index]

        # Add total row
        tot_row = {}
        for m in months:
            m_str = m.strftime('%Y-%m')
            tot_row[(m_str, "Sum")] = total[m]
            tot_row[(m_str, "Percent")] = ""
            tot_row[(m_str, "Comment")] = ""
        combined.loc["Current period total"] = tot_row

        # Now attach the top-level multiindex for row => (field, value_label)
        combined.index = pd.MultiIndex.from_product([[field], combined.index],
                                                    names=["Field Name", "Value Label"])
        frames.append(combined)

    if not frames:
        return pd.DataFrame()

    final = pd.concat(frames)
    return final

def flatten_dataframe(df):
    if isinstance(df.columns, pd.MultiIndex):
        df = df.reset_index()
        df.columns = [" ".join(map(str, col)).strip() if isinstance(col, tuple) else col 
                      for col in df.columns.values]
    return df

def load_report_data(file_path, date1, date2):
    df_data = pd.read_excel(file_path, sheet_name="Data")
    df_data["filemonth_dt"] = pd.to_datetime(df_data["filemonth_dt"])

    wb = load_workbook(file_path, data_only=True)
    # summary
    if "Summary" in wb.sheetnames:
        summary_df = pd.read_excel(file_path, sheet_name="Summary")
    else:
        summary_df = generate_summary_df(df_data, date1, date2)

    # Value Distribution with monthly comment columns
    if "Value Distribution" in wb.sheetnames:
        val_dist_df = pd.read_excel(file_path, sheet_name="Value Distribution")
    else:
        val_dist_df = generate_dist_with_comments(df_data, "value_dist", date1)

    # Population Comparison with monthly comment columns
    if "Population Comparison" in wb.sheetnames:
        pop_comp_df = pd.read_excel(file_path, sheet_name="Population Comparison")
    else:
        pop_comp_df = generate_dist_with_comments(df_data, "pop_comp", date1)

    return df_data, summary_df, val_dist_df, pop_comp_df

#############################################
# Functions to Cache Previous Comments
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
                        # skip if raw is "" or "nan"
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
            [str(item) for item in x
             if pd.notnull(item) and str(item).strip().lower() != "nan"]
        )
    ).reset_index()
    grouped = grouped.rename(columns={"Comment": f"comment_{target_month}"})
    return grouped

#############################################
# Function to Preserve Existing Summary Comments
#############################################

def preserve_summary_comments(input_file_path, summary_df):
    try:
        existing = pd.read_excel(input_file_path, sheet_name="Summary")
        comment_dict = {}
        approval_dict = {}
        if "Comment" in existing.columns:
            comment_dict = existing.set_index("Field Name")["Comment"].to_dict()
        if "Approval Comments" in existing.columns:
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

    # Cache previous comments immediately
    prev_comments_df = get_cached_previous_comments(folder)
    if prev_comments_df.empty:
        prev_comments_df = cache_previous_comments(folder)

    st.write("Cached previous comments:")
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
        # Value Distribution Grid
        ##############################
        st.subheader("Value Distribution")
        val_fields = st.session_state.value_dist_df["Field Name"].unique().tolist()
        if not val_fields:
            st.warning("No Value Distribution data available.")
        else:
            active_val = st.session_state.active_field if st.session_state.active_field in val_fields else val_fields[0]
            selected_val_field = st.selectbox("Select Field (Value Dist)",
                                              val_fields,
                                              index=val_fields.index(active_val),
                                              key="val_field_select")
            st.session_state.active_field = selected_val_field

            filtered_val = st.session_state.value_dist_df[
                st.session_state.value_dist_df["Field Name"] == selected_val_field
            ].copy()

            # We might have monthly columns like: "2025-01 Sum", "2025-01 Percent", "2025-01 Comment", ...
            # We'll do an example for editing. We'll skip the aggregator approach for now.

            gb_val = GridOptionsBuilder.from_dataframe(filtered_val)
            gb_val.configure_default_column(
                editable=True,
                cellStyle={'white-space': 'normal','line-height':'1.2em','width':150}
            )

            # We'll detect columns that end with "Comment" and make them editable
            for c in filtered_val.columns:
                if c.endswith("Comment"):
                    gb_val.configure_column(c, editable=True, width=200)
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
            val_res = AgGrid(
                filtered_val,
                gridOptions=val_opts,
                update_mode=GridUpdateMode.VALUE_CHANGED,
                data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
                key="val_grid",
                height=val_height,
                use_container_width=True
            )
            st.session_state.value_dist_df = pd.DataFrame(val_res["data"]).copy()

        ##############################
        # Population Comparison Grid
        ##############################
        st.subheader("Population Comparison")
        pop_fields = st.session_state.pop_comp_df["Field Name"].unique().tolist()
        if not pop_fields:
            st.warning("No Population Comparison data available.")
        else:
            active_pop = st.session_state.active_field if st.session_state.active_field in pop_fields else pop_fields[0]
            selected_pop_field = st.selectbox("Select Field (Pop Comp)",
                                              pop_fields,
                                              index=pop_fields.index(active_pop) if active_pop in pop_fields else 0,
                                              key="pop_field_select")
            st.session_state.active_field = selected_pop_field

            filtered_pop = st.session_state.pop_comp_df[
                st.session_state.pop_comp_df["Field Name"] == selected_pop_field
            ].copy()

            gb_pop = GridOptionsBuilder.from_dataframe(filtered_pop)
            gb_pop.configure_default_column(
                editable=True,
                cellStyle={'white-space': 'normal','line-height':'1.2em','width':150}
            )

            for c in filtered_pop.columns:
                if c.endswith("Comment"):
                    gb_pop.configure_column(c, editable=True, width=200)
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
            pop_res = AgGrid(
                filtered_pop,
                gridOptions=pop_opts,
                update_mode=GridUpdateMode.VALUE_CHANGED,
                data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
                key="pop_grid",
                height=pop_height,
                use_container_width=True
            )
            st.session_state.pop_comp_df = pd.DataFrame(pop_res["data"]).copy()

        # aggregator skipping "nan"
        def aggregate_current_comments():
            """
            We'll parse each monthly comment column in Value Dist & Pop Comp,
            skipping lines if comment is 'nan' or empty, then join them into
            the Summary's 'Comment' column.
            """
            sum_df = st.session_state.summary_df.copy()
            # For each field in summary, gather lines from the monthly comment columns
            for field in sum_df["Field Name"].unique():
                notes = []

                # 1) From Value Dist
                vdist = st.session_state.value_dist_df[
                    st.session_state.value_dist_df["Field Name"] == field
                ]
                # monthly columns end with "Comment"
                cmt_cols = [c for c in vdist.columns if c.endswith("Comment")]
                for _, row in vdist.iterrows():
                    val_label = row.get("Value Label", "")
                    for c in cmt_cols:
                        raw_comment = str(row.get(c, "")).strip()
                        if raw_comment.lower() == "nan" or raw_comment == "":
                            continue
                        # e.g. "2025-01 Comment"
                        # we can parse out the month
                        month_name = c.replace(" Comment","").strip()
                        # Build line
                        line = f"{val_label} ({month_name}) - {raw_comment}".strip(" -")
                        notes.append(line)

                # 2) From Pop Comp
                pcomp = st.session_state.pop_comp_df[
                    st.session_state.pop_comp_df["Field Name"] == field
                ]
                cmt_cols2 = [c for c in pcomp.columns if c.endswith("Comment")]
                for _, row in pcomp.iterrows():
                    val_label = row.get("Value Label", "")
                    for c in cmt_cols2:
                        raw_comment = str(row.get(c, "")).strip()
                        if raw_comment.lower() == "nan" or raw_comment == "":
                            continue
                        month_name = c.replace(" Comment","").strip()
                        line = f"{val_label} ({month_name}) - {raw_comment}".strip(" -")
                        notes.append(line)

                aggregated_note = "\n".join(notes).strip()
                if aggregated_note:
                    sum_df.loc[sum_df["Field Name"] == field, "Comment"] = aggregated_note

            st.session_state.summary_df = sum_df

        aggregate_current_comments()

        # Reorder summary columns
        sum_df = st.session_state.summary_df.copy()
        cols = list(sum_df.columns)
        cols.remove("Approval Comments")
        cols.remove("Comment")
        new_order = ["Field Name"] + cols[1:] + ["Comment", "Approval Comments"]
        sum_df = sum_df[new_order]

        st.subheader("Summary")
        gb_sum = GridOptionsBuilder.from_dataframe(sum_df)
        gb_sum.configure_default_column(
            editable=False,
            cellStyle={'white-space':'normal','line-height':'1.2em','width':150}
        )
        gb_sum.configure_column("Approval Comments", editable=True, width=250, minWidth=100, maxWidth=300)
        gb_sum.configure_column("Comment", editable=False, width=250, minWidth=100, maxWidth=300)
        for c in sum_df.columns:
            if c not in ["Field Name", "Comment", "Approval Comments"]:
                if "Change" in c:
                    gb_sum.configure_column(
                        c,
                        type=["numericColumn"],
                        valueFormatter="(params.value != null ? params.value.toFixed(2)+'%' : '')",
                        width=150, minWidth=100, maxWidth=200
                    )
                else:
                    gb_sum.configure_column(
                        c,
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

        sum_res = AgGrid(
            sum_df,
            gridOptions=sum_opts,
            update_mode=GridUpdateMode.VALUE_CHANGED,
            data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
            key="sum_grid",
            height=sum_height,
            use_container_width=True
        )
        st.session_state.summary_df = pd.DataFrame(sum_res["data"]).copy()

        # Now in-place update
        try:
            with pd.ExcelWriter(st.session_state.input_file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                # Summary
                export_sum = st.session_state.summary_df.copy().reset_index(drop=True)
                export_sum.to_excel(writer, index=False, sheet_name="Summary")
                summary_sheet = writer.sheets["Summary"]

                # Insert prev-month comment in "Month-to-Month Diff (date1)" column
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
                            if prev_comment_str.lower() != "nan" and prev_comment_str != "":
                                cell = summary_sheet.cell(row=idx+2, column=d1_col_index)
                                com_obj = Comment(prev_comment_str, "Prev")
                                com_obj.visible = True
                                cell.comment = com_obj

                # Value Distribution
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
                            if prev_comment_str.lower() != "nan" and prev_comment_str != "":
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

                # Population Comparison
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
                            if prev_comment_str.lower() != "nan" and prev_comment_str != "":
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