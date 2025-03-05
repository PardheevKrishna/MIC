import streamlit as st
import pandas as pd
import os
import datetime
import re
from io import BytesIO
from dateutil.relativedelta import relativedelta

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
        phrases = ["1\\)   F6CF Loan - Both Pop, Diff Values", "2\\)   CF Loan - Prior Null, Current Pop", "3\\)   CF Loan - Prior Pop, Current Null"]
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
    final_df = pd.concat(result_frames)
    final_df.columns = pd.MultiIndex.from_tuples(final_df.columns)
    return final_df

def main():
    st.set_page_config(page_title="Official FRY14M Field Analysis Summary", layout="centered", initial_sidebar_state="expanded")
    st.sidebar.title("File & Date Selection")
    predefined_mapping = {"BDCOM": ["bdcom_report.xlsx", "bdcom_data.xlsb"], "wifinsa": ["wifinsa_report.xlsx", "wifinsa_data.xlsb"]}
    category = st.sidebar.selectbox("Select Category", ["BDCOM", "wifinsa"])
    folder_files = os.listdir('.')
    available_files = [f for f in predefined_mapping.get(category, []) if f in folder_files]
    if not available_files:
        st.sidebar.error(f"No Excel files for {category} found in the current folder as per mapping.")
        return
    selected_file = st.sidebar.selectbox("Select an Excel File", available_files)
    selected_date = st.sidebar.date_input("Select Date for Date1", datetime.date(2025, 1, 1))
    date1 = datetime.datetime.combine(selected_date, datetime.datetime.min.time())
    date2 = date1 - relativedelta(months=1)
    if st.sidebar.button("Generate Report"):
        generate_report(selected_file, date1, date2)

def generate_report(file_path, date1, date2):
    st.title("Official FRY14M Field Analysis Summary Report")
    if not file_path.lower().endswith('.xlsx'):
        engine = get_excel_engine(file_path)
        if engine:
            df_data = pd.read_excel(file_path, sheet_name="Data", engine=engine)
        else:
            df_data = pd.read_excel(file_path, sheet_name="Data")
        temp = BytesIO()
        with pd.ExcelWriter(temp, engine='openpyxl') as writer:
            df_data.to_excel(writer, index=False, sheet_name="Data")
        temp.seek(0)
        df_data = pd.read_excel(temp, sheet_name="Data", engine='openpyxl')
    else:
        df_data = pd.read_excel(file_path, sheet_name="Data")
    required_cols = {"analysis_type", "filemonth_dt", "field_name", "value_label", "value_records"}
    if not required_cols.issubset(df_data.columns):
        st.error(f"'Data' sheet must contain columns: {required_cols}. Found: {list(df_data.columns)}")
        return
    try:
        df_data["filemonth_dt"] = pd.to_datetime(df_data["filemonth_dt"])
    except Exception as e:
        st.error(f"Error parsing 'filemonth_dt': {e}")
        return
    st.write(f"**Selected File:** {file_path}")
    st.write(f"**Date1:** {date1.strftime('%Y-%m-%d')} | **Date2:** {date2.strftime('%Y-%m-%d')}")
    summary_df = generate_summary_df(df_data, date1, date2)
    st.subheader("Summary Results")
    st.dataframe(summary_df, use_container_width=True)
    st.subheader("Value Distribution")
    value_dist_df = generate_distribution_df(df_data, "value_dist", date1)
    st.dataframe(value_dist_df, use_container_width=True)
    st.subheader("Population Comparison")
    pop_comp_df = generate_distribution_df(df_data, "pop_comp", date1)
    st.dataframe(pop_comp_df, use_container_width=True)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        summary_df.to_excel(writer, index=False, sheet_name="Summary")
        value_dist_df.to_excel(writer, sheet_name="Value Distribution")
        pop_comp_df.to_excel(writer, sheet_name="Population Comparison")
    processed_data = output.getvalue()
    st.download_button(label="Download Report as Excel", data=processed_data, file_name="FRY14M_Field_Analysis_Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    main()