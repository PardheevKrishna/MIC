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

##############################
# Helper Functions
##############################

def compute_grid_height(df, row_height=40, header_height=80):
    n = len(df)
    return header_height + (min(n, 30) * row_height)

def get_excel_engine(file_path):
    # For .xlsx files we use openpyxl by default.
    return None

def generate_summary_df(df_data, date1, date2):
    fields = sorted(df_data["field_name"].unique())
    rows = []
    for field in fields:
        mask_miss_d1 = (
            (df_data['analysis_type'] == 'value_dist') &
            (df_data['field_name'] == field) &
            (df_data['filemonth_dt'] == date1) &
            (df_data['value_label'].str.contains("Missing", case=False, na=False))
        )
        missing_d1 = df_data.loc[mask_miss_d1, 'value_records'].sum()

        mask_miss_d2 = (
            (df_data['analysis_type'] == 'value_dist') &
            (df_data['field_name'] == field) &
            (df_data['filemonth_dt'] == date2) &
            (df_data['value_label'].str.contains("Missing", case=False, na=False))
        )
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

        mask_pop_d1 = (
            (df_data['analysis_type'] == 'pop_comp') &
            (df_data['field_name'] == field) &
            (df_data['filemonth_dt'] == date1) &
            (df_data['value_label'].apply(contains_phrase))
        )
        pop_d1 = df_data.loc[mask_pop_d1, 'value_records'].sum()

        mask_pop_d2 = (
            (df_data['analysis_type'] == 'pop_comp') &
            (df_data['field_name'] == field) &
            (df_data['filemonth_dt'] == date2) &
            (df_data['value_label'].apply(contains_phrase))
        )
        pop_d2 = df_data.loc[mask_pop_d2, 'value_records'].sum()

        rows.append([field, missing_d1, missing_d2, pop_d1, pop_d2])

    df = pd.DataFrame(rows, columns=[
        "Field Name",
        f"Missing Values ({date1.strftime('%Y-%m-%d')})",
        f"Missing Values ({date2.strftime('%Y-%m-%d')})",
        f"Month-to-Month Diff ({date1.strftime('%Y-%m-%d')})",
        f"Month-to-Month Diff ({date2.strftime('%Y-%m-%d')})"
    ])

    # Calculate percentage changes:
    m1 = f"Missing Values ({date1.strftime('%Y-%m-%d')})"
    m2 = f"Missing Values ({date2.strftime('%Y-%m-%d')})"
    d1 = f"Month-to-Month Diff ({date1.strftime('%Y-%m-%d')})"
    d2 = f"Month-to-Month Diff ({date2.strftime('%Y-%m-%d')})"

    df["Missing % Change"] = df.apply(lambda r: ((r[m1]-r[m2]) / r[m2] * 100) if r[m2]!=0 else None, axis=1)
    df["Month-to-Month % Change"] = df.apply(lambda r: ((r[d1]-r[d2]) / r[d2] * 100) if r[d2]!=0 else None, axis=1)
    
    new_order = [
        "Field Name",
        m1,
        m2,
        "Missing % Change",
        d1,
        d2,
        "Month-to-Month % Change"
    ]
    df = df[new_order]
    df["Comment"] = ""  # To store aggregated current comments
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
        df.columns = [" ".join(map(str, col)).strip() if isinstance(col, tuple) else col 
                      for col in df.columns.values]
    return df

def load_report_data(file_path, date1, date2):
    df_data = pd.read_excel(file_path, sheet_name="Data")
    df_data["filemonth_dt"] = pd.to_datetime(df_data["filemonth_dt"])
    summary_df = generate_summary_df(df_data, date1, date2)
    val_dist_df = generate_distribution_df(df_data, "value_dist", date1)
    pop_comp_df = generate_distribution_df(df_data, "pop_comp", date1)
    return df_data, summary_df, val_dist_df, pop_comp_df

#########################################
# Functions for Caching Previous Comments
#########################################

def cache_previous_comments(current_folder):
    """
    Scan folder "previous/<current_folder>" for .xlsx files,
    extract cell comments from the Summary sheet (from the column starting with "Month-to-Month Diff"),
    and save a CSV with columns: Field Name, Month, Comment.
    """
    data = []
    prev_folder = os.path.join(os.getcwd(), "previous", current_folder)
    if not os.path.exists(prev_folder):
        st.warning("Previous months folder not found in " + prev_folder)
        return pd.DataFrame(columns=["Field Name", "Month", "Comment"])
    for file in os.listdir(prev_folder):
        if file.lower().endswith('.xlsx'):
            file_path = os.path.join(prev_folder, file)
            m = re.search(r'(\d{4}-\d{2})', file)
            month_year = m.group(1) if m else "unknown"
            try:
                wb = load_workbook(file_path, data_only=True)
            except Exception as e:
                st.error(f"Error opening previous file {file}: {e}")
                continue
            if "Summary" not in wb.sheetnames:
                continue
            ws = wb["Summary"]
            header = [cell.value for cell in ws[1]]
            col_index = None
            for i, col_name in enumerate(header, start=1):
                if col_name and str(col_name).startswith("Month-to-Month Diff"):
                    col_index = i
                    break
            if col_index is None:
                continue
            for row in ws.iter_rows(min_row=2):
                field_cell = row[0]
                if field_cell.value:
                    field_name = str(field_cell.value).strip()
                    cell = row[col_index - 1]  # adjust for 0-index
                    comment_text = cell.comment.text if cell.comment else ""
                    data.append({"Field Name": field_name, "Month": month_year, "Comment": comment_text})
    df = pd.DataFrame(data)
    df.to_csv("previous_comments.csv", index=False)
    return df

def get_cached_previous_comments(current_folder):
    if os.path.exists("previous_comments.csv"):
        df = pd.read_csv("previous_comments.csv")
    else:
        df = cache_previous_comments(current_folder)
    return df

def pivot_previous_comments(df):
    """
    Pivot the cached previous comments so that for each Field Name, 
    each Month becomes a column (named "comment_<Month>").
    """
    if df.empty:
        return pd.DataFrame()
    grouped = df.groupby(["Field Name", "Month"])["Comment"].apply(lambda x: "\n".join(x)).unstack(fill_value="")
    grouped = grouped.rename(columns=lambda x: f"comment_{x}")
    grouped.reset_index(inplace=True)
    return grouped

#########################################
# Main Streamlit App
#########################################

def main():
    st.sidebar.title("File & Date Selection")
    folder = st.sidebar.selectbox("Select Folder", ["BDCOM", "WFHMSA", "BCards"])
    folder_path = os.path.join(os.getcwd(), folder)
    st.sidebar.write(f"Folder path: {folder_path}")
    if not os.path.exists(folder_path):
        st.sidebar.error(f"Folder '{folder}' not found.")
        return
    # Cache previous comments immediately at startup.
    prev_comments_df = cache_previous_comments(folder)
    st.session_state.prev_comments_df = prev_comments_df
    st.write("Cached previous comments CSV created (if previous files exist).")

    # Only .xlsx files are supported for in-place update.
    all_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.xlsx')]
    if not all_files:
        st.sidebar.error(f"No Excel files found in folder '{folder}'.")
        return
    selected_file = st.sidebar.selectbox("Select an Excel File", all_files)
    input_file_path = os.path.join(folder_path, selected_file)
    selected_date = st.sidebar.date_input("Select Date for Date1", datetime.date(2025, 1, 1))
    date1 = datetime.datetime.combine(selected_date, datetime.datetime.min.time())
    date2 = date1 - relativedelta(months=1)
    
    if st.sidebar.button("Generate Report"):
        df_data, summary_df, val_dist_df, pop_comp_df = load_report_data(input_file_path, date1, date2)
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
        # (Assume AgGrid code for displaying and updating Value Distribution, Population Comparison, and Summary)
        # For brevity, we assume that st.session_state.value_dist_df and st.session_state.pop_comp_df
        # have been updated by the grids.
        # ---------------------------
        
        # ---------------------------
        # Aggregate current grid comments into Summary
        def aggregate_comments_into_summary():
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

        aggregate_comments_into_summary()
        
        # ---------------------------
        # Retrieve and pivot previous comments from the cached CSV
        prev_df = get_cached_previous_comments(st.session_state.folder)
        pivot_prev = pivot_previous_comments(prev_df)
        if not pivot_prev.empty:
            st.session_state.summary_df = st.session_state.summary_df.merge(pivot_prev, on="Field Name", how="left")
            prev_months = sorted(pivot_prev.columns.drop("Field Name").tolist())
            selected_prev = st.selectbox("Select Previous Month", options=prev_months, format_func=lambda x: x.replace("comment_", ""), key="prev_month")
        else:
            selected_prev = None

        # ---------------------------
        # Display the updated Summary Grid (including previous comment columns)
        st.subheader("Summary")
        sum_df = st.session_state.summary_df.copy()
        gb_sum = GridOptionsBuilder.from_dataframe(sum_df)
        gb_sum.configure_default_column(editable=False,
                                        cellStyle={'white-space': 'normal', 'line-height': '1.2em', 'width': 150})
        gb_sum.configure_column("Comment", editable=False, width=250, minWidth=100, maxWidth=300)
        for col in sum_df.columns:
            if col not in ["Field Name", "Comment"] and not col.startswith("comment_"):
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
        
        # (Optionally, if a previous month is selected, update the detail grids with those comments.)
        if selected_prev:
            prev_col = selected_prev  # e.g., "comment_2024-12"
            vd = st.session_state.value_dist_df.copy()
            vd["Comment"] = vd["Field Name"].apply(lambda x: st.session_state.summary_df.loc[st.session_state.summary_df["Field Name"]==x, prev_col].values[0] if prev_col in st.session_state.summary_df.columns else "")
            st.session_state.value_dist_df = vd
            pc = st.session_state.pop_comp_df.copy()
            pc["Comment"] = pc["Field Name"].apply(lambda x: st.session_state.summary_df.loc[st.session_state.summary_df["Field Name"]==x, prev_col].values[0] if prev_col in st.session_state.summary_df.columns else "")
            st.session_state.pop_comp_df = pc
        
        # ---------------------------
        # In-Place Update of the Input Excel File
        try:
            with pd.ExcelWriter(st.session_state.input_file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                # --- Summary Sheet ---
                export_sum = st.session_state.summary_df.copy().reset_index(drop=True)
                sum_comments = export_sum["Comment"]
                export_sum.drop(columns=["Comment"], inplace=True, errors="ignore")
                export_sum.to_excel(writer, index=False, sheet_name="Summary")
                summary_sheet = writer.sheets["Summary"]
                d1_col_name = f"Month-to-Month Diff ({st.session_state.date1.strftime('%Y-%m-%d')})"
                sum_cols = export_sum.columns.tolist()
                try:
                    d1_col_index = sum_cols.index(d1_col_name) + 1
                except ValueError:
                    d1_col_index = export_sum.shape[1]
                for idx, comm in enumerate(sum_comments):
                    excel_row = idx + 2
                    if str(comm).strip():
                        cell = summary_sheet.cell(row=excel_row, column=d1_col_index)
                        com_obj = Comment(str(comm), "User")
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
                    agg_note = st.session_state.summary_df.loc[st.session_state.summary_df["Field Name"] == field, "Comment"].values
                    if len(agg_note) > 0 and agg_note[0].strip():
                        if vd_sum_col_index is not None:
                            cell = vd_sheet.cell(row=excel_row, column=vd_sum_col_index)
                            com_obj = Comment(agg_note[0], "User")
                            com_obj.visible = True
                            cell.comment = com_obj
                        if vd_percent_col_index is not None:
                            cell = vd_sheet.cell(row=excel_row, column=vd_percent_col_index)
                            com_obj = Comment(agg_note[0], "User")
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
                    agg_note = st.session_state.summary_df.loc[st.session_state.summary_df["Field Name"] == field, "Comment"].values
                    if len(agg_note) > 0 and agg_note[0].strip():
                        if pop_sum_col_index is not None:
                            cell = pop_sheet.cell(row=excel_row, column=pop_sum_col_index)
                            com_obj = Comment(agg_note[0], "User")
                            com_obj.visible = True
                            cell.comment = com_obj
                        if pop_percent_col_index is not None:
                            cell = pop_sheet.cell(row=excel_row, column=pop_percent_col_index)
                            com_obj = Comment(agg_note[0], "User")
                            com_obj.visible = True
                            cell.comment = com_obj

                writer.save()
            st.success("The input Excel file has been updated successfully!")
        except Exception as e:
            st.error(f"Error updating the Excel file: {e}")

if __name__ == "__main__":
    main()