import streamlit as st
st.set_page_config(page_title="Official FRY14M Field Analysis Summary", layout="centered", initial_sidebar_state="expanded")

import pandas as pd
import os
import datetime
import re
from io import BytesIO
from dateutil.relativedelta import relativedelta
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode

def get_excel_engine(file_path):
    if file_path.lower().endswith('.xlsb'):
        return 'pyxlsb'
    return None

def generate_summary_df(df_data, date1, date2):
    fields = sorted(df_data["field_name"].unique())
    summary_data = []
    for field in fields:
        mask_missing_date1 = ((df_data['analysis_type'] == 'value_dist') & (df_data['field_name'] == field) & (df_data['filemonth_dt'] == date1) & (df_data['value_label'].str.contains("Missing", case=False, na=False)))
        missing_sum_d1 = df_data.loc[mask_missing_date1, 'value_records'].sum()
        mask_missing_date2 = ((df_data['analysis_type'] == 'value_dist') & (df_data['field_name'] == field) & (df_data['filemonth_dt'] == date2) & (df_data['value_label'].str.contains("Missing", case=False, na=False)))
        missing_sum_d2 = df_data.loc[mask_missing_date2, 'value_records'].sum()
        phrases = ["1\\)   CF Loan - Both Pop, Diff Values", "2\\)   CF Loan - Prior Null, Current Pop", "3\\)   CF Loan - Prior Pop, Current Null"]
        def contains_phrase(x):
            for pat in phrases:
                if re.search(pat, x):
                    return True
            return False
        mask_m2m_date1 = ((df_data['analysis_type'] == 'pop_comp') & (df_data['field_name'] == field) & (df_data['filemonth_dt'] == date1) & (df_data['value_label'].apply(lambda x: contains_phrase(x))))
        m2m_sum_d1 = df_data.loc[mask_m2m_date1, 'value_records'].sum()
        mask_m2m_date2 = ((df_data['analysis_type'] == 'pop_comp') & (df_data['field_name'] == field) & (df_data['filemonth_dt'] == date2) & (df_data['value_label'].apply(lambda x: contains_phrase(x))))
        m2m_sum_d2 = df_data.loc[mask_m2m_date2, 'value_records'].sum()
        summary_data.append([field, missing_sum_d1, missing_sum_d2, m2m_sum_d1, m2m_sum_d2])
    summary_df = pd.DataFrame(summary_data, columns=["Field Name", f"Missing Values ({date1.strftime('%Y-%m-%d')})", f"Missing Values ({date2.strftime('%Y-%m-%d')})", f"Month-to-Month Diff ({date1.strftime('%Y-%m-%d')})", f"Month-to-Month Diff ({date2.strftime('%Y-%m-%d')})"])
    return summary_df

def generate_distribution_df(df, analysis_type, date1):
    months = [(date1 - relativedelta(months=i)).replace(day=1) for i in range(0, 12)]
    months = sorted(months, reverse=True)
    df_filtered = df[df['analysis_type'] == analysis_type].copy()
    df_filtered['month'] = df_filtered['filemonth_dt'].apply(lambda d: d.replace(day=1))
    df_filtered = df_filtered[df_filtered['month'].isin(months)]
    grouped = df_filtered.groupby(['field_name', 'value_label', 'month'])['value_records'].sum().reset_index()
    if grouped.empty:
        st.warning(f"No data found for analysis type: {analysis_type}")
        return pd.DataFrame()
    pivot = grouped.pivot_table(index=['field_name', 'value_label'], columns='month', values='value_records', fill_value=0)
    pivot = pivot.reindex(columns=months, fill_value=0)
    result_frames = []
    for field, sub_df in pivot.groupby(level=0):
        sub_df = sub_df.droplevel(0)
        total = sub_df.sum(axis=0)
        percent_df = sub_df.div(total, axis=1).multiply(100).round(2).fillna(0)
        data = {}
        for m in months:
            m_str = m.strftime('%Y-%m')
            data[(m_str, "Sum")] = sub_df[m]
            data[(m_str, "Percent")] = percent_df[m]
        temp_df = pd.DataFrame(data)
        total_row = {}
        for m in months:
            m_str = m.strftime('%Y-%m')
            total_row[(m_str, "Sum")] = total[m]
            total_row[(m_str, "Percent")] = ""
        temp_df.loc["Current period total"] = total_row
        temp_df.index = pd.MultiIndex.from_product([[field], temp_df.index], names=["Field Name", "Value Label"])
        result_frames.append(temp_df)
    if not result_frames:
        st.warning(f"No distribution data available for analysis type: {analysis_type}")
        return pd.DataFrame()
    final_df = pd.concat(result_frames)
    final_df.columns = pd.MultiIndex.from_tuples(final_df.columns)
    return final_df

def flatten_dataframe(df):
    if isinstance(df.columns, pd.MultiIndex):
        df = df.reset_index()
        df.columns = [' '.join(map(str, col)).strip() if isinstance(col, tuple) else col for col in df.columns.values]
    return df

def load_report_data(file_path, date1, date2):
    if file_path.lower().endswith('.xlsb'):
        df_data = pd.read_excel(file_path, sheet_name="Data", engine='pyxlsb')
    else:
        df_data = pd.read_excel(file_path, sheet_name="Data")
    df_data["filemonth_dt"] = pd.to_datetime(df_data["filemonth_dt"])
    summary_df = generate_summary_df(df_data, date1, date2)
    value_dist_df = generate_distribution_df(df_data, "value_dist", date1)
    pop_comp_df = generate_distribution_df(df_data, "pop_comp", date1)
    return df_data, summary_df, value_dist_df, pop_comp_df

st.write("Working Directory:", os.getcwd())

def main():
    st.sidebar.title("File & Date Selection")
    folder = st.sidebar.selectbox("Select Folder", ["BDCOM", "WFHMSA"])
    folder_path = os.path.join(os.getcwd(), folder)
    st.sidebar.write(f"Folder path: {folder_path}")
    if not os.path.exists(folder_path):
        st.sidebar.error(f"Folder '{folder}' not found in the working directory.")
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
        df_data, summary_df, value_dist_df, pop_comp_df = load_report_data(os.path.join(folder_path, selected_file), date1, date2)
        st.session_state.report_generated = True
        st.session_state.selected_file = selected_file
        st.session_state.folder = folder
        st.session_state.date1 = date1
        st.session_state.date2 = date2
        st.session_state.df_data = df_data
        st.session_state.summary_df = summary_df
        st.session_state.value_dist_df = flatten_dataframe(value_dist_df.copy())
        st.session_state.pop_comp_df = flatten_dataframe(pop_comp_df.copy())
    if st.session_state.get("report_generated", False):
        st.title("Official FRY14M Field Analysis Summary Report")
        st.write(f"**Folder:** {st.session_state.folder}")
        st.write(f"**Selected File:** {st.session_state.selected_file}")
        st.write(f"**Date1:** {st.session_state.date1.strftime('%Y-%m-%d')} | **Date2:** {st.session_state.date2.strftime('%Y-%m-%d')}")
        
        # --- Summary Grid with Linking ---
        summary_grid = st.session_state.summary_df.copy()
        if "Comments" not in summary_grid.columns:
            summary_grid["Comments"] = ""
        gb_summary = GridOptionsBuilder.from_dataframe(summary_grid)
        gb_summary.configure_default_column(editable=True, singleClickEdit=True)
        gb_summary.configure_selection('single', use_checkbox=False)
        gridOptions_summary = gb_summary.build()
        st.subheader("Summary Results")
        summary_response = AgGrid(
            summary_grid,
            gridOptions=gridOptions_summary,
            update_mode=GridUpdateMode.VALUE_CHANGED,
            data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
            key="summary_grid"
        )
        st.session_state.summary_df = pd.DataFrame(summary_response["data"])
        selected_summary = summary_response.get("selectedRows", [])
        if selected_summary:
            field_link = selected_summary[0].get("Field Name")
            unique_fields_value = st.session_state.value_dist_df["Field Name"].unique().tolist()
            unique_fields_pop = st.session_state.pop_comp_df["Field Name"].unique().tolist()
            if field_link in unique_fields_value:
                st.session_state.current_field_value = unique_fields_value.index(field_link)
            if field_link in unique_fields_pop:
                st.session_state.current_field_pop = unique_fields_pop.index(field_link)
        
        # --- Value Distribution Grid with Linking ---
        st.subheader("Value Distribution")
        unique_fields_value = st.session_state.value_dist_df["Field Name"].unique().tolist()
        if "current_field_value" not in st.session_state:
            st.session_state.current_field_value = 0
        cols_value = st.columns([1,1])
        if cols_value[0].button("←", key="prev_value"):
            st.session_state.current_field_value = (st.session_state.current_field_value - 1) % len(unique_fields_value)
        if cols_value[1].button("→", key="next_value"):
            st.session_state.current_field_value = (st.session_state.current_field_value + 1) % len(unique_fields_value)
        selected_field_value = st.selectbox("Select Field", unique_fields_value, index=st.session_state.current_field_value, key="select_value")
        st.session_state.current_field_value = unique_fields_value.index(selected_field_value)
        filtered_value_dist = st.session_state.value_dist_df[st.session_state.value_dist_df["Field Name"] == selected_field_value]
        if "Comments" not in filtered_value_dist.columns:
            filtered_value_dist["Comments"] = ""
        gb_value = GridOptionsBuilder.from_dataframe(filtered_value_dist)
        gb_value.configure_default_column(editable=True, singleClickEdit=True)
        gb_value.configure_selection('single', use_checkbox=False)
        gridOptions_value = gb_value.build()
        value_response = AgGrid(
            filtered_value_dist,
            gridOptions=gridOptions_value,
            update_mode=GridUpdateMode.VALUE_CHANGED,
            data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
            key="value_dist_grid"
        )
        temp_value = pd.DataFrame(value_response["data"])
        st.session_state.value_dist_df.update(temp_value)
        selected_value = value_response.get("selectedRows", [])
        if selected_value:
            field_sel = selected_value[0].get("Field Name")
            unique_fields_pop = st.session_state.pop_comp_df["Field Name"].unique().tolist()
            if field_sel in unique_fields_pop:
                st.session_state.current_field_pop = unique_fields_pop.index(field_sel)
        
        # --- Population Comparison Grid with Linking and SQL Logic Display ---
        st.subheader("Population Comparison")
        unique_fields_pop = st.session_state.pop_comp_df["Field Name"].unique().tolist()
        if "current_field_pop" not in st.session_state:
            st.session_state.current_field_pop = 0
        cols_pop = st.columns([1,1])
        if cols_pop[0].button("←", key="prev_pop"):
            st.session_state.current_field_pop = (st.session_state.current_field_pop - 1) % len(unique_fields_pop)
        if cols_pop[1].button("→", key="next_pop"):
            st.session_state.current_field_pop = (st.session_state.current_field_pop + 1) % len(unique_fields_pop)
        selected_field_pop = st.selectbox("Select Field", unique_fields_pop, index=st.session_state.current_field_pop, key="select_pop")
        st.session_state.current_field_pop = unique_fields_pop.index(selected_field_pop)
        filtered_pop_comp = st.session_state.pop_comp_df[st.session_state.pop_comp_df["Field Name"] == selected_field_pop]
        if "Comments" not in filtered_pop_comp.columns:
            filtered_pop_comp["Comments"] = ""
        gb_pop = GridOptionsBuilder.from_dataframe(filtered_pop_comp)
        gb_pop.configure_default_column(editable=True, singleClickEdit=True)
        gb_pop.configure_selection('single', use_checkbox=False)
        gridOptions_pop = gb_pop.build()
        pop_response = AgGrid(
            filtered_pop_comp,
            gridOptions=gridOptions_pop,
            update_mode=GridUpdateMode.VALUE_CHANGED,
            data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
            key="pop_comp_grid"
        )
        temp_pop = pd.DataFrame(pop_response["data"])
        st.session_state.pop_comp_df.update(temp_pop)
        selected_pop = pop_response.get("selectedRows", [])
        if selected_pop:
            field_sel = selected_pop[0].get("Field Name")
            value_label = selected_pop[0].get("Value Label")
            if value_label != "Current period total":
                df_data = st.session_state.df_data
                sql_logic_vals = df_data[(df_data["field_name"] == field_sel) & (df_data["value_label"] == value_label)]["value_sql_logic"].unique()
                st.text_area("Value SQL Logic", value="\n".join(sql_logic_vals) if sql_logic_vals.size > 0 else "No SQL Logic found")
        
        # --- Excel Download ---
        summary_updated = st.session_state.summary_df
        value_updated = st.session_state.value_dist_df
        pop_updated = st.session_state.pop_comp_df
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            summary_updated.to_excel(writer, index=False, sheet_name="Summary")
            value_updated.to_excel(writer, index=False, sheet_name="Value Distribution")
            pop_updated.to_excel(writer, index=False, sheet_name="Population Comparison")
        processed_data = output.getvalue()
        st.download_button(label="Download Report as Excel", data=processed_data, file_name="FRY14M_Field_Analysis_Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    main()