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

# Custom CSS for header wrapping and full container usage
st.markdown("""
    <style>
        .ag-header-cell-label {
            white-space: pre-wrap !important;
            text-align: center;
            line-height: 1.2em;
        }
        .ag-root-wrapper {
            margin: 0 !important;
            padding: 0 !important;
        }
    </style>
""", unsafe_allow_html=True)

def compute_grid_height(df, row_height=40, header_height=80):
    """Dynamically compute grid height based on the number of rows."""
    n = len(df)
    if n == 0:
        return header_height + 20
    return min(n, 30) * row_height + header_height

def get_excel_engine(file_path):
    """Return 'pyxlsb' engine if the file ends with .xlsb, otherwise None."""
    return 'pyxlsb' if file_path.lower().endswith('.xlsb') else None

def generate_summary_df(df_data, date1, date2):
    """Generate the Summary DataFrame comparing two months for each field."""
    fields = sorted(df_data["field_name"].unique())
    rows = []
    for field in fields:
        # Missing for date1
        mask_miss_d1 = (
            (df_data['analysis_type'] == 'value_dist') &
            (df_data['field_name'] == field) &
            (df_data['filemonth_dt'] == date1) &
            (df_data['value_label'].str.contains("Missing", case=False, na=False))
        )
        missing_d1 = df_data.loc[mask_miss_d1, 'value_records'].sum()

        # Missing for date2
        mask_miss_d2 = (
            (df_data['analysis_type'] == 'value_dist') &
            (df_data['field_name'] == field) &
            (df_data['filemonth_dt'] == date2) &
            (df_data['value_label'].str.contains("Missing", case=False, na=False))
        )
        missing_d2 = df_data.loc[mask_miss_d2, 'value_records'].sum()

        # Example phrases used in pop_comp
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

        # Pop comp for date1
        mask_pop_d1 = (
            (df_data['analysis_type'] == 'pop_comp') &
            (df_data['field_name'] == field) &
            (df_data['filemonth_dt'] == date1) &
            (df_data['value_label'].apply(contains_phrase))
        )
        pop_d1 = df_data.loc[mask_pop_d1, 'value_records'].sum()

        # Pop comp for date2
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
    m1 = f"Missing Values ({date1.strftime('%Y-%m-%d')})"
    m2 = f"Missing Values ({date2.strftime('%Y-%m-%d')})"
    d1 = f"Month-to-Month Diff ({date1.strftime('%Y-%m-%d')})"
    d2 = f"Month-to-Month Diff ({date2.strftime('%Y-%m-%d')})"

    # Calculate percentage change
    df["Missing % Change"] = df.apply(lambda r: ((r[m1] - r[m2]) / r[m2] * 100) if r[m2] != 0 else None, axis=1)
    df["Month-to-Month % Change"] = df.apply(lambda r: ((r[d1] - r[d2]) / r[d2] * 100) if r[d2] != 0 else None, axis=1)

    # Reorder columns so percentage columns follow their related value columns
    new_order = [
        "Field Name",
        f"Missing Values ({date1.strftime('%Y-%m-%d')})",
        "Missing % Change",
        f"Missing Values ({date2.strftime('%Y-%m-%d')})",
        f"Month-to-Month Diff ({date1.strftime('%Y-%m-%d')})",
        "Month-to-Month % Change",
        f"Month-to-Month Diff ({date2.strftime('%Y-%m-%d')})"
    ]
    df = df[new_order]

    # Add one comment column at the end.
    df["Comment"] = ""
    return df

def generate_distribution_df(df, analysis_type, date1):
    """
    Generate a distribution table (value_dist or pop_comp) with 12 months of data.
    This returns a pivot with (Sum, Percent) columns for each month.
    """
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

        # Add a "Current period total" row
        tot_row = {}
        for m in months:
            m_str = m.strftime('%Y-%m')
            tot_row[(m_str, "Sum")] = total[m]
            tot_row[(m_str, "Percent")] = ""
        tmp.loc["Current period total"] = tot_row

        # Rebuild a MultiIndex for the final DataFrame
        tmp.index = pd.MultiIndex.from_product([[field], tmp.index], names=["Field Name", "Value Label"])
        frames.append(tmp)

    if not frames:
        return pd.DataFrame()

    final = pd.concat(frames)
    final.columns = pd.MultiIndex.from_tuples(final.columns)
    return final

def flatten_dataframe(df):
    """
    Flatten a multi-index DataFrame so that it can be displayed easily in st.aggrid.
    """
    if isinstance(df.columns, pd.MultiIndex):
        df = df.reset_index()
        df.columns = [
            " ".join(map(str, col)).strip() if isinstance(col, tuple) else col
            for col in df.columns.values
        ]
    return df

def load_report_data(file_path, date1, date2):
    """Load the report data from Excel and generate the summary, value_dist, and pop_comp DataFrames."""
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

st.write("Working Directory:", os.getcwd())

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
    selected_date = st.sidebar.date_input("Select Date for Date1", datetime.date(2025, 1, 1))
    date1 = datetime.datetime.combine(selected_date, datetime.datetime.min.time())
    date2 = date1 - relativedelta(months=1)
    
    # Generate Report button
    if st.sidebar.button("Generate Report"):
        df_data, summary_df, val_dist_df, pop_comp_df = load_report_data(
            os.path.join(folder_path, selected_file), date1, date2
        )
        st.session_state.df_data = df_data
        st.session_state.summary_df = summary_df
        st.session_state.value_dist_df = flatten_dataframe(val_dist_df.copy())
        st.session_state.pop_comp_df = flatten_dataframe(pop_comp_df.copy())
        st.session_state.selected_file = selected_file
        st.session_state.folder = folder
        st.session_state.date1 = date1
        st.session_state.date2 = date2
        st.session_state.active_field = None

    # If we have data in session state, render the UI
    if st.session_state.get("df_data") is not None:
        st.title("FRY14M Field Analysis Summary Report")
        st.write(f"**Folder:** {st.session_state.folder}")
        st.write(f"**File:** {st.session_state.selected_file}")
        st.write(f"**Date1:** {st.session_state.date1.strftime('%Y-%m-%d')} | **Date2:** {st.session_state.date2.strftime('%Y-%m-%d')}")

        # --------------------------------------------------------------------
        # Summary Grid
        # --------------------------------------------------------------------
        sum_df = st.session_state.summary_df.copy()
        if "Comment" not in sum_df.columns:
            sum_df["Comment"] = ""

        gb_sum = GridOptionsBuilder.from_dataframe(sum_df)
        gb_sum.configure_default_column(
            editable=False,
            cellStyle={'white-space': 'normal', 'line-height': '1.2em', "width": 150}
        )
        # Enable editing for the "Comment" column
        gb_sum.configure_column("Comment", editable=True, width=150, minWidth=100, maxWidth=200)

        # Configure numeric columns
        for col in sum_df.columns:
            if col not in ["Field Name", "Comment"]:
                if "Change" in col:
                    gb_sum.configure_column(
                        col,
                        type=["numericColumn"],
                        valueFormatter="(params.value != null ? params.value.toFixed(2) + '%' : '')",
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

        # Adjust header formatting
        for c in sum_opts["columnDefs"]:
            if "headerName" in c:
                c["headerName"] = "\n".join(c["headerName"].split())
                c["width"] = 150
                c["minWidth"] = 100
                c["maxWidth"] = 200

        gb_sum.configure_selection("single", use_checkbox=False)
        sum_opts["rowSelection"] = "single"
        sum_opts["pagination"] = False
        sum_opts["rowHeight"] = 40
        sum_opts["headerHeight"] = 80

        sum_height = compute_grid_height(sum_df, row_height=40, header_height=80)
        st.subheader("Summary")
        sum_res = AgGrid(
            sum_df,
            gridOptions=sum_opts,
            update_mode=GridUpdateMode.VALUE_CHANGED,
            data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
            key="sum_grid",
            height=sum_height,
            use_container_width=True
        )
        # Updated summary data after user edits
        updated_sum = pd.DataFrame(sum_res["data"])
        st.session_state.summary_df = updated_sum

        # Propagate summary "Comment" to value_dist_df and pop_comp_df by matching Field Name
        for field_name in updated_sum["Field Name"].unique():
            comment_for_field = updated_sum.loc[updated_sum["Field Name"] == field_name, "Comment"].iloc[0]
            st.session_state.value_dist_df.loc[
                st.session_state.value_dist_df["Field Name"] == field_name, "Comment"
            ] = comment_for_field
            st.session_state.pop_comp_df.loc[
                st.session_state.pop_comp_df["Field Name"] == field_name, "Comment"
            ] = comment_for_field

        sel_sum = sum_res.get("selectedRows", [])
        if sel_sum:
            st.session_state.active_field = sel_sum[0].get("Field Name")

        # --------------------------------------------------------------------
        # Value Distribution Grid
        # --------------------------------------------------------------------
        st.subheader("Value Distribution")
        val_fields = st.session_state.value_dist_df["Field Name"].unique().tolist()
        # Active field logic
        active_val = st.session_state.active_field if st.session_state.get("active_field") in val_fields else val_fields[0]

        selected_val_field = st.selectbox("Field (Value Dist)", val_fields,
                                          index=val_fields.index(active_val),
                                          key="val_select")
        st.session_state.active_field = selected_val_field

        filtered_val = st.session_state.value_dist_df[
            st.session_state.value_dist_df["Field Name"] == selected_val_field
        ]
        if "Comment" not in filtered_val.columns:
            filtered_val["Comment"] = ""

        gb_val = GridOptionsBuilder.from_dataframe(filtered_val)
        gb_val.configure_default_column(
            editable=False,
            cellStyle={'white-space': 'normal', 'line-height': '1.2em', "width": 150}
        )
        gb_val.configure_column("Comment", editable=True, width=150, minWidth=100, maxWidth=200)
        gb_val.configure_selection("single", use_checkbox=False)

        val_opts = gb_val.build()
        if isinstance(val_opts, list):
            val_opts = {"columnDefs": val_opts}

        for c in val_opts["columnDefs"]:
            if "headerName" in c:
                c["headerName"] = "\n".join(c["headerName"].split())
                c["width"] = 150
                c["minWidth"] = 100
                c["maxWidth"] = 200

        val_opts["rowSelection"] = "single"
        val_opts["pagination"] = False
        val_opts["rowHeight"] = 40
        val_opts["headerHeight"] = 80

        val_height = compute_grid_height(filtered_val, row_height=40, header_height=80)
        val_res = AgGrid(
            filtered_val,
            gridOptions=val_opts,
            update_mode=GridUpdateMode.VALUE_CHANGED,
            data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
            key="val_grid",
            height=val_height,
            use_container_width=True
        )

        # Capture the selected row in Value Distribution
        val_selected = val_res.get("selectedRows", [])
        if not val_selected and not filtered_val.empty:
            preselect_val_label_dist = filtered_val.iloc[0]["Value Label"]
        else:
            preselect_val_label_dist = (
                val_selected[0]["Value Label"] if val_selected and "Value Label" in val_selected[0] else None
            )

        # Update st.session_state.value_dist_df with new comment changes
        updated_val = pd.DataFrame(val_res["data"])
        st.session_state.value_dist_df.update(updated_val)

        # --------------------------------------------------------------------
        # Population Comparison Grid
        # --------------------------------------------------------------------
        st.subheader("Population Comparison")
        pop_fields = st.session_state.pop_comp_df["Field Name"].unique().tolist()
        active_pop = st.session_state.active_field if st.session_state.get("active_field") in pop_fields else pop_fields[0]

        selected_pop_field = st.selectbox("Field (Pop Comp)", pop_fields,
                                          index=pop_fields.index(active_pop),
                                          key="pop_select")
        st.session_state.active_field = selected_pop_field

        filtered_pop = st.session_state.pop_comp_df[
            st.session_state.pop_comp_df["Field Name"] == selected_pop_field
        ].copy()
        if "Comment" not in filtered_pop.columns:
            filtered_pop["Comment"] = ""

        gb_pop = GridOptionsBuilder.from_dataframe(filtered_pop)
        gb_pop.configure_default_column(
            editable=False,
            cellStyle={'white-space': 'normal', 'line-height': '1.2em', "width": 150}
        )
        gb_pop.configure_column("Comment", editable=True, width=150, minWidth=100, maxWidth=200)
        gb_pop.configure_selection("single", use_checkbox=False)

        pop_opts = gb_pop.build()
        if isinstance(pop_opts, list):
            pop_opts = {"columnDefs": pop_opts}

        for c in pop_opts["columnDefs"]:
            if "headerName" in c:
                c["headerName"] = "\n".join(c["headerName"].split())
                c["width"] = 150
                c["minWidth"] = 100
                c["maxWidth"] = 200

        pop_opts["rowSelection"] = "single"
        pop_opts["pagination"] = False
        pop_opts["rowHeight"] = 40
        pop_opts["headerHeight"] = 80

        pop_height = compute_grid_height(filtered_pop, row_height=40, header_height=80)
        pop_res = AgGrid(
            filtered_pop,
            gridOptions=pop_opts,
            update_mode=GridUpdateMode.VALUE_CHANGED,
            data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
            key="pop_grid",
            height=pop_height,
            use_container_width=True
        )
        pop_selected = pop_res.get("selectedRows", [])
        if not pop_selected and not filtered_pop.empty:
            preselect_val_label_pop = filtered_pop.iloc[0]["Value Label"]
        else:
            preselect_val_label_pop = (
                pop_selected[0]["Value Label"] if pop_selected and "Value Label" in pop_selected[0] else None
            )

        # Update st.session_state.pop_comp_df with new comment changes
        updated_pop = pd.DataFrame(pop_res["data"])
        st.session_state.pop_comp_df.update(updated_pop)

        # --------------------------------------------------------------------
        # View SQL Logic (Population Comparison)
        # --------------------------------------------------------------------
        st.subheader("View SQL Logic (Population Comparison)")
        pop_orig = st.session_state.df_data[st.session_state.df_data["analysis_type"] == "pop_comp"]
        pop_orig_field = pop_orig[pop_orig["field_name"] == selected_pop_field]
        pop_val_labels = pop_orig_field["value_label"].dropna().unique().tolist()

        if not pop_val_labels:
            st.write("No Value Labels available for SQL Logic (Pop Comp).")
        else:
            default_pop_val = (
                preselect_val_label_pop
                if (preselect_val_label_pop in pop_val_labels)
                else (pop_val_labels[0] if pop_val_labels else None)
            )
            sel_pop_val_label = st.selectbox(
                "Select Value Label (Pop Comp)",
                pop_val_labels,
                index=pop_val_labels.index(default_pop_val) if default_pop_val else 0,
                key="pop_sql_val_label"
            )
            months = [(st.session_state.date1 - relativedelta(months=i)).replace(day=1) for i in range(12)]
            months = sorted(months, reverse=True)
            month_options = [m.strftime("%Y-%m") for m in months]
            sel_pop_month = st.selectbox("Select Month (Pop Comp)", month_options, key="pop_sql_month")

            if st.button("Show SQL Logic (Pop Comp)"):
                matches = pop_orig_field[
                    (pop_orig_field["value_label"] == sel_pop_val_label) &
                    (pop_orig_field["filemonth_dt"].dt.strftime("%Y-%m") == sel_pop_month)
                ]
                sql_vals = matches["value_sql_logic"].dropna().unique()
                if sql_vals.size > 0:
                    st.text_area("Value SQL Logic (Pop Comp)", "\n".join(sql_vals), height=150)
                else:
                    st.text_area("Value SQL Logic (Pop Comp)", "No SQL Logic found", height=150)

        # --------------------------------------------------------------------
        # View SQL Logic (Value Distribution)
        # --------------------------------------------------------------------
        st.subheader("View SQL Logic (Value Distribution)")
        val_orig = st.session_state.df_data[st.session_state.df_data["analysis_type"] == "value_dist"]
        val_orig_field = val_orig[val_orig["field_name"] == selected_val_field]
        val_val_labels = val_orig_field["value_label"].dropna().unique().tolist()

        if not val_val_labels:
            st.write("No Value Labels available for SQL Logic (Value Dist).")
        else:
            default_val_val = (
                preselect_val_label_dist
                if (preselect_val_label_dist in val_val_labels)
                else (val_val_labels[0] if val_val_labels else None)
            )
            sel_val_val_label = st.selectbox(
                "Select Value Label (Value Dist)",
                val_val_labels,
                index=val_val_labels.index(default_val_val) if default_val_val else 0,
                key="val_sql_val_label"
            )
            months_val = [(st.session_state.date1 - relativedelta(months=i)).replace(day=1) for i in range(12)]
            months_val = sorted(months_val, reverse=True)
            month_options_val = [m.strftime("%Y-%m") for m in months_val]
            sel_val_month = st.selectbox("Select Month (Value Dist)", month_options_val, key="val_sql_month")

            if st.button("Show SQL Logic (Value Dist)"):
                matches_val = val_orig_field[
                    (val_orig_field["value_label"] == sel_val_val_label) &
                    (val_orig_field["filemonth_dt"].dt.strftime("%Y-%m") == sel_val_month)
                ]
                sql_vals_val = matches_val["value_sql_logic"].dropna().unique()
                if sql_vals_val.size > 0:
                    st.text_area("Value SQL Logic (Value Dist)", "\n".join(sql_vals_val), height=150)
                else:
                    st.text_area("Value SQL Logic (Value Dist)", "No SQL Logic found", height=150)

        # --------------------------------------------------------------------
        # Excel Download with Cell Comments for ALL Sheets
        # --------------------------------------------------------------------
        out_buf = BytesIO()
        with pd.ExcelWriter(out_buf, engine='openpyxl') as writer:
            # 1) Summary
            export_sum = st.session_state.summary_df.copy()
            sum_comments = export_sum["Comment"].tolist()
            export_sum.drop(columns=["Comment"], inplace=True, errors="ignore")
            export_sum.to_excel(writer, index=False, sheet_name="Summary")

            # 2) Value Distribution
            export_val_dist = st.session_state.value_dist_df.copy()
            val_dist_comments = export_val_dist["Comment"].tolist()
            export_val_dist.drop(columns=["Comment"], inplace=True, errors="ignore")
            export_val_dist.to_excel(writer, index=False, sheet_name="Value Distribution")

            # 3) Population Comparison
            export_pop_comp = st.session_state.pop_comp_df.copy()
            pop_comp_comments = export_pop_comp["Comment"].tolist()
            export_pop_comp.drop(columns=["Comment"], inplace=True, errors="ignore")
            export_pop_comp.to_excel(writer, index=False, sheet_name="Population Comparison")

            # Now add each comment as a cell note in the last column of each sheet
            workbook = writer.book

            # Summary comments
            sheet_sum = writer.sheets["Summary"]
            sum_col_count = export_sum.shape[1]  # last column index (since 'Comment' was dropped)
            for row_idx, comm_text in enumerate(sum_comments, start=2):
                if str(comm_text).strip():
                    cell = sheet_sum.cell(row=row_idx, column=sum_col_count)
                    cell.comment = Comment(str(comm_text), "User")

            # Value Distribution comments
            sheet_val = writer.sheets["Value Distribution"]
            val_col_count = export_val_dist.shape[1]
            for row_idx, comm_text in enumerate(val_dist_comments, start=2):
                if str(comm_text).strip():
                    cell = sheet_val.cell(row=row_idx, column=val_col_count)
                    cell.comment = Comment(str(comm_text), "User")

            # Population Comparison comments
            sheet_pop = writer.sheets["Population Comparison"]
            pop_col_count = export_pop_comp.shape[1]
            for row_idx, comm_text in enumerate(pop_comp_comments, start=2):
                if str(comm_text).strip():
                    cell = sheet_pop.cell(row=row_idx, column=pop_col_count)
                    cell.comment = Comment(str(comm_text), "User")

        st.download_button(
            "Download Report as Excel",
            data=out_buf.getvalue(),
            file_name="FRY14M_Field_Analysis_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()