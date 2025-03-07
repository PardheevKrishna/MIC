import streamlit as st
st.set_page_config(page_title="Final FRY14M Field Analysis", layout="centered", initial_sidebar_state="expanded")

import pandas as pd
import os
import datetime
from io import BytesIO
from dateutil.relativedelta import relativedelta
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode

def get_excel_engine(file_path):
    if file_path.lower().endswith('.xlsb'):
        return 'pyxlsb'
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
    return pd.DataFrame(rows, columns=[
        "Field Name",
        f"Missing Values ({date1.strftime('%Y-%m-%d')})",
        f"Missing Values ({date2.strftime('%Y-%m-%d')})",
        f"Month-to-Month Diff ({date1.strftime('%Y-%m-%d')})",
        f"Month-to-Month Diff ({date2.strftime('%Y-%m-%d')})"
    ])

def generate_distribution_df(df, analysis_type, date1):
    months = [(date1 - relativedelta(months=i)).replace(day=1) for i in range(12)]
    months = sorted(months, reverse=True)
    sub = df[df['analysis_type'] == analysis_type].copy()
    sub['month'] = sub['filemonth_dt'].apply(lambda d: d.replace(day=1))
    sub = sub[sub['month'].isin(months)]
    grouped = sub.groupby(['field_name', 'value_label', 'month'])['value_records'].sum().reset_index()
    if grouped.empty:
        return pd.DataFrame()
    pivot = grouped.pivot_table(
        index=['field_name', 'value_label'], 
        columns='month', 
        values='value_records', 
        fill_value=0
    )
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
        df.columns = [
            " ".join(map(str, col)).strip() if isinstance(col, tuple) else col 
            for col in df.columns.values
        ]
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
    folder = st.sidebar.selectbox("Select Folder", ["BDCOM", "WFHMSA"])
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
        st.session_state.val_field_index = 0
        st.session_state.pop_field_index = 0

    if st.session_state.get("df_data") is not None:
        st.title("FRY14M Field Analysis Summary Report")
        st.write(f"**Folder:** {st.session_state.folder}")
        st.write(f"**File:** {st.session_state.selected_file}")
        st.write(f"**Date1:** {st.session_state.date1.strftime('%Y-%m-%d')} | **Date2:** {st.session_state.date2.strftime('%Y-%m-%d')}")
        
        # --- Summary Grid (Read-only except Comments) ---
        sum_df = st.session_state.summary_df.copy()
        if "Comments" not in sum_df.columns:
            sum_df["Comments"] = ""
        gb_sum = GridOptionsBuilder.from_dataframe(sum_df)
        gb_sum.configure_default_column(editable=False)
        gb_sum.configure_column("Comments", editable=True)
        gb_sum.configure_selection("single", use_checkbox=False)
        sum_opts = gb_sum.build()
        sum_opts["rowSelection"] = "single"
        st.subheader("Summary")
        sum_res = AgGrid(
            sum_df,
            gridOptions=sum_opts,
            update_mode=GridUpdateMode.SELECTION_CHANGED,
            data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
            key="sum_grid"
        )
        sel_sum = sum_res.get("selectedRows", [])
        if sel_sum:
            link_field = sel_sum[0].get("Field Name")
            val_fields = st.session_state.value_dist_df["Field Name"].unique().tolist()
            pop_fields = st.session_state.pop_comp_df["Field Name"].unique().tolist()
            if link_field in val_fields:
                st.session_state.val_field_index = val_fields.index(link_field)
            if link_field in pop_fields:
                st.session_state.pop_field_index = pop_fields.index(link_field)
        
        # --- Value Distribution Grid ---
        st.subheader("Value Distribution")
        val_fields = st.session_state.value_dist_df["Field Name"].unique().tolist()
        col1, col2, col3 = st.columns([1,1,1])
        if col1.button("←", key="prev_val"):
            st.session_state.val_field_index = (st.session_state.val_field_index - 1) % len(val_fields)
        if col3.button("→", key="next_val"):
            st.session_state.val_field_index = (st.session_state.val_field_index + 1) % len(val_fields)
        sel_val_field = st.selectbox("Field (Value Dist)", val_fields, index=st.session_state.val_field_index, key="val_select")
        st.session_state.val_field_index = val_fields.index(sel_val_field)
        filtered_val = st.session_state.value_dist_df[st.session_state.value_dist_df["Field Name"] == sel_val_field]
        if "Comments" not in filtered_val.columns:
            filtered_val["Comments"] = ""
        gb_val = GridOptionsBuilder.from_dataframe(filtered_val)
        gb_val.configure_default_column(editable=False)
        gb_val.configure_column("Comments", editable=True)
        gb_val.configure_selection("single", use_checkbox=False)
        val_opts = gb_val.build()
        val_opts["rowSelection"] = "single"
        val_res = AgGrid(
            filtered_val,
            gridOptions=val_opts,
            update_mode=GridUpdateMode.SELECTION_CHANGED,
            data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
            key="val_grid"
        )
        sel_val = val_res.get("selectedRows", [])
        if sel_val:
            field_from_val = sel_val[0].get("Field Name")
            pop_fields = st.session_state.pop_comp_df["Field Name"].unique().tolist()
            if field_from_val in pop_fields:
                st.session_state.pop_field_index = pop_fields.index(field_from_val)
        
        # --- Population Comparison Grid ---
        st.subheader("Population Comparison")
        pop_fields = st.session_state.pop_comp_df["Field Name"].unique().tolist()
        col4, col5, col6 = st.columns([1,1,1])
        if col4.button("←", key="prev_pop"):
            st.session_state.pop_field_index = (st.session_state.pop_field_index - 1) % len(pop_fields)
        if col6.button("→", key="next_pop"):
            st.session_state.pop_field_index = (st.session_state.pop_field_index + 1) % len(pop_fields)
        sel_pop_field = st.selectbox("Field (Pop Comp)", pop_fields, index=st.session_state.pop_field_index, key="pop_select")
        st.session_state.pop_field_index = pop_fields.index(sel_pop_field)
        filtered_pop = st.session_state.pop_comp_df[st.session_state.pop_comp_df["Field Name"] == sel_pop_field].copy()
        if "Comments" not in filtered_pop.columns:
            filtered_pop["Comments"] = ""
        gb_pop = GridOptionsBuilder.from_dataframe(filtered_pop)
        gb_pop.configure_default_column(editable=False)
        gb_pop.configure_column("Comments", editable=True)
        gb_pop.configure_selection("single", use_checkbox=False)
        pop_opts = gb_pop.build()
        pop_opts["rowSelection"] = "single"
        pop_res = AgGrid(
            filtered_pop,
            gridOptions=pop_opts,
            update_mode=GridUpdateMode.SELECTION_CHANGED,
            data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
            key="pop_grid"
        )
        sel_pop = pop_res.get("selectedRows", [])
        # --- View SQL Logic Section ---
        if sel_pop:
            pop_field = sel_pop[0].get("Field Name")
            pop_value = sel_pop[0].get("Value Label")
            if pop_value != "Current period total":
                # Filter original data for matching pop_comp row
                orig = st.session_state.df_data
                match = orig[
                    (orig["analysis_type"] == "pop_comp") &
                    (orig["field_name"] == pop_field) &
                    (orig["value_label"] == pop_value)
                ]
                sql_vals = match["value_sql_logic"].unique()
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