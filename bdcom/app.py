import streamlit as st
st.set_page_config(page_title="Final FRY14M Field Analysis", layout="wide", initial_sidebar_state="expanded")

import pandas as pd
import os
import datetime
import re
from io import BytesIO
from dateutil.relativedelta import relativedelta
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode

# Custom CSS for header wrapping and full width
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

def compute_grid_height(df, row_height=40, header_height=35):
    n = len(df)
    if n == 0:
        return header_height + 20
    # Use 40px per row and show exactly n rows if n < 30, else 30 rows.
    return (n if n < 30 else 30)*row_height + header_height

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
    # Compute percentage change columns:
    m1 = f"Missing Values ({date1.strftime('%Y-%m-%d')})"
    m2 = f"Missing Values ({date2.strftime('%Y-%m-%d')})"
    d1 = f"Month-to-Month Diff ({date1.strftime('%Y-%m-%d')})"
    d2 = f"Month-to-Month Diff ({date2.strftime('%Y-%m-%d')})"
    df["Missing % Change"] = df.apply(lambda r: ((r[m1] - r[m2]) / r[m2] * 100) if r[m2] != 0 else None, axis=1)
    df["Month-to-Month % Change"] = df.apply(lambda r: ((r[d1] - r[d2]) / r[d2] * 100) if r[d2] != 0 else None, axis=1)
    # Reorder columns: insert percentage columns next to their related values.
    new_order = [
        "Field Name",
        f"Missing Values ({date1.strftime('%Y-%m-%d')})",
        "Missing % Change",
        f"Missing Values ({date2.strftime('%Y-%m-%d')})",
        f"Month-to-Month Diff ({date1.strftime('%Y-%m-%d')})",
        "Month-to-Month % Change",
        f"Month-to-Month Diff ({date2.strftime('%Y-%m-%d')})"
    ]
    return df[new_order]

def generate_distribution_df(df, analysis_type, date1):
    months = [(date1 - relativedelta(months=i)).replace(day=1) for i in range(12)]
    months = sorted(months, reverse=True)
    sub = df[df['analysis_type'] == analysis_type].copy()
    sub['month'] = sub['filemonth_dt'].apply(lambda d: d.replace(day=1))
    sub = sub[sub['month'].isin(months)]
    grouped = sub.groupby(['field_name', 'value_label', 'month'])['value_records'].sum().reset_index()
    if grouped.empty:
        return pd.DataFrame()
    pivot = grouped.pivot_table(index=['field_name', 'value_label'], columns='month', values='value_records', fill_value=0)
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

    if st.session_state.get("df_data") is not None:
        st.title("FRY14M Field Analysis Summary Report")
        st.write(f"**Folder:** {st.session_state.folder}")
        st.write(f"**File:** {st.session_state.selected_file}")
        st.write(f"**Date1:** {st.session_state.date1.strftime('%Y-%m-%d')} | **Date2:** {st.session_state.date2.strftime('%Y-%m-%d')}")
        
        # --- Summary Grid ---
        sum_df = st.session_state.summary_df.copy()
        if "Comments" not in sum_df.columns:
            sum_df["Comments"] = ""
        gb_sum = GridOptionsBuilder.from_dataframe(sum_df)
        # Set cell style and row height via gridOptions property "rowHeight"
        gb_sum.configure_default_column(
            editable=False,
            cellStyle={'white-space': 'normal', 'line-height': '1.2em', "width": 25}
        )
        for col in sum_df.columns:
            if col not in ["Field Name", "Comments"]:
                if "Change" in col:
                    gb_sum.configure_column(col, valueFormatter="(params.value != null ? params.value.toFixed(2) + '%' : '')", width=25)
                else:
                    gb_sum.configure_column(col, valueFormatter="(params.value != null ? params.value.toLocaleString('en-US') : '')", width=25)
        sum_opts = gb_sum.build()
        if isinstance(sum_opts, list):
            sum_opts = {"columnDefs": sum_opts}
        for c in sum_opts["columnDefs"]:
            if "headerName" in c:
                c["headerName"] = "\n".join(c["headerName"].split())
                c["width"] = 25
        gb_sum.configure_selection("single", use_checkbox=False)
        # Use the already built sum_opts; do not call build() again.
        sum_opts["rowSelection"] = "single"
        sum_opts["pagination"] = False
        sum_opts["rowHeight"] = 40
        sum_height = compute_grid_height(sum_df, row_height=40, header_height=35)
        st.subheader("Summary")
        sum_res = AgGrid(
            sum_df,
            gridOptions=sum_opts,
            update_mode=GridUpdateMode.SELECTION_CHANGED,
            data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
            key="sum_grid",
            height=sum_height,
            use_container_width=True
        )
        sel_sum = sum_res.get("selectedRows", [])
        if sel_sum:
            st.session_state.active_field = sel_sum[0].get("Field Name")
        
        # --- Value Distribution Grid ---
        st.subheader("Value Distribution")
        val_fields = st.session_state.value_dist_df["Field Name"].unique().tolist()
        active_val = st.session_state.active_field if st.session_state.get("active_field") in val_fields else val_fields[0]
        selected_val_field = st.selectbox("Field (Value Dist)", val_fields, index=val_fields.index(active_val), key="val_select")
        st.session_state.active_field = selected_val_field
        filtered_val = st.session_state.value_dist_df[st.session_state.value_dist_df["Field Name"] == selected_val_field]
        if "Comments" not in filtered_val.columns:
            filtered_val["Comments"] = ""
        gb_val = GridOptionsBuilder.from_dataframe(filtered_val)
        gb_val.configure_default_column(
            editable=False,
            cellStyle={'white-space': 'normal', 'line-height': '1.2em', "width": 25}
        )
        gb_val.configure_column("Comments", editable=True, width=25)
        gb_val.configure_selection("single", use_checkbox=False)
        val_opts = gb_val.build()
        if isinstance(val_opts, list):
            val_opts = {"columnDefs": val_opts}
        for c in val_opts["columnDefs"]:
            if "headerName" in c:
                c["headerName"] = "\n".join(c["headerName"].split())
                c["width"] = 25
        val_opts["rowSelection"] = "single"
        val_opts["pagination"] = False
        val_opts["rowHeight"] = 40
        val_height = compute_grid_height(filtered_val, row_height=40, header_height=35)
        AgGrid(
            filtered_val,
            gridOptions=val_opts,
            update_mode=GridUpdateMode.SELECTION_CHANGED,
            data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
            key="val_grid",
            height=val_height,
            use_container_width=True
        )
        
        # --- Population Comparison Grid ---
        st.subheader("Population Comparison")
        pop_fields = st.session_state.pop_comp_df["Field Name"].unique().tolist()
        active_pop = st.session_state.active_field if st.session_state.get("active_field") in pop_fields else pop_fields[0]
        selected_pop_field = st.selectbox("Field (Pop Comp)", pop_fields, index=pop_fields.index(active_pop), key="pop_select")
        st.session_state.active_field = selected_pop_field
        filtered_pop = st.session_state.pop_comp_df[st.session_state.pop_comp_df["Field Name"] == selected_pop_field].copy()
        if "Comments" not in filtered_pop.columns:
            filtered_pop["Comments"] = ""
        gb_pop = GridOptionsBuilder.from_dataframe(filtered_pop)
        gb_pop.configure_default_column(
            editable=False,
            cellStyle={'white-space': 'normal', 'line-height': '1.2em', "width": 25}
        )
        gb_pop.configure_column("Comments", editable=True, width=25)
        gb_pop.configure_selection("single", use_checkbox=False)
        pop_opts = gb_pop.build()
        if isinstance(pop_opts, list):
            pop_opts = {"columnDefs": pop_opts}
        for c in pop_opts["columnDefs"]:
            if "headerName" in c:
                c["headerName"] = "\n".join(c["headerName"].split())
                c["width"] = 25
        pop_opts["rowSelection"] = "single"
        pop_opts["pagination"] = False
        pop_opts["rowHeight"] = 40
        pop_height = compute_grid_height(filtered_pop, row_height=40, header_height=35)
        pop_res = AgGrid(
            filtered_pop,
            gridOptions=pop_opts,
            update_mode=GridUpdateMode.SELECTION_CHANGED,
            data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
            key="pop_grid",
            height=pop_height,
            use_container_width=True
        )
        sel_pop = pop_res.get("selectedRows", [])
        
        # --- View SQL Logic Section (via Dropdowns) ---
        st.subheader("View SQL Logic")
        orig = st.session_state.df_data
        pop_orig = orig[orig["analysis_type"] == "pop_comp"]
        pop_orig_field = pop_orig[pop_orig["field_name"] == selected_pop_field]
        val_labels = pop_orig_field["value_label"].dropna().unique().tolist()
        if not val_labels:
            st.write("No Value Labels available for SQL Logic.")
        else:
            sel_val_label = st.selectbox("Select Value Label", val_labels, key="sql_val_label")
            months = [(st.session_state.date1 - relativedelta(months=i)).replace(day=1) for i in range(12)]
            months = sorted(months, reverse=True)
            month_options = [m.strftime("%Y-%m") for m in months]
            sel_month = st.selectbox("Select Month", month_options, key="sql_month")
            if st.button("Show SQL Logic"):
                matches = pop_orig_field[
                    (pop_orig_field["value_label"] == sel_val_label) &
                    (pop_orig_field["filemonth_dt"].dt.strftime("%Y-%m") == sel_month)
                ]
                sql_vals = matches["value_sql_logic"].dropna().unique()
                if sql_vals.size > 0:
                    st.text_area("Value SQL Logic", "\n".join(sql_vals), height=150)
                else:
                    st.text_area("Value SQL Logic", "No SQL Logic found", height=150)
        
        # --- Excel Download ---
        out_buf = BytesIO()
        with pd.ExcelWriter(out_buf, engine='openpyxl') as writer:
            st.session_state.summary_df.to_excel(writer, index=False, sheet_name="Summary")
            st.session_state.value_dist_df.to_excel(writer, index=False, sheet_name="Value Distribution")
            st.session_state.pop_comp_df.to_excel(writer, index=False, sheet_name="Population Comparison")
        st.download_button(
            "Download Report as Excel",
            data=out_buf.getvalue(),
            file_name="FRY14M_Field_Analysis_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()