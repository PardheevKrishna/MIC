import streamlit as st
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
        mask_missing_date1 = (
            (df_data['analysis_type'] == 'value_dist') &
            (df_data['field_name'] == field) &
            (df_data['filemonth_dt'] == date1) &
            (df_data['value_label'].str.contains("Missing", case=False, na=False))
        )
        missing_sum_d1 = df_data.loc[mask_missing_date1, 'value_records'].sum()
        
        mask_missing_date2 = (
            (df_data['analysis_type'] == 'value_dist') &
            (df_data['field_name'] == field) &
            (df_data['filemonth_dt'] == date2) &
            (df_data['value_label'].str.contains("Missing", case=False, na=False))
        )
        missing_sum_d2 = df_data.loc[mask_missing_date2, 'value_records'].sum()
        
        phrases = [
            "1\\)   CF Loan - Both Pop, Diff Values",
            "2\\)   CF Loan - Prior Null, Current Pop",
            "3\\)   CF Loan - Prior Pop, Current Null"
        ]
        
        def contains_phrase(x):
            for pat in phrases:
                if re.search(pat, x):
                    return True
            return False
        
        mask_m2m_date1 = (
            (df_data['analysis_type'] == 'pop_comp') &
            (df_data['field_name'] == field) &
            (df_data['filemonth_dt'] == date1) &
            (df_data['value_label'].apply(lambda x: contains_phrase(x)))
        )
        m2m_sum_d1 = df_data.loc[mask_m2m_date1, 'value_records'].sum()
        
        mask_m2m_date2 = (
            (df_data['analysis_type'] == 'pop_comp') &
            (df_data['field_name'] == field) &
            (df_data['filemonth_dt'] == date2) &
            (df_data['value_label'].apply(lambda x: contains_phrase(x)))
        )
        m2m_sum_d2 = df_data.loc[mask_m2m_date2, 'value_records'].sum()
        
        summary_data.append([
            field,
            missing_sum_d1,
            missing_sum_d2,
            m2m_sum_d1,
            m2m_sum_d2
        ])
    
    summary_df = pd.DataFrame(
        summary_data,
        columns=[
            "Field Name",
            f"Missing Values ({date1.strftime('%Y-%m-%d')})",
            f"Missing Values ({date2.strftime('%Y-%m-%d')})",
            f"Month-to-Month Diff ({date1.strftime('%Y-%m-%d')})",
            f"Month-to-Month Diff ({date2.strftime('%Y-%m-%d')})"
        ]
    )
    
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


def flatten_dataframe(df):
    if isinstance(df.columns, pd.MultiIndex):
        df = df.reset_index()
        df.columns = [' '.join(map(str, col)).strip() if isinstance(col, tuple) else col for col in df.columns.values]
    return df


def load_comments_for_grid(df, csv_file, key_cols):
    if os.path.exists(csv_file):
        comments_df = pd.read_csv(csv_file)
        df = df.merge(comments_df, on=key_cols, how='left', suffixes=('', '_saved'))
        df["Comments"] = df["Comments_saved"].fillna(df["Comments"])
        df.drop(columns=["Comments_saved"], inplace=True)
    return df


def save_comments_for_grid(df, csv_file, key_cols):
    comments_df = df[key_cols + ["Comments"]]
    comments_df.to_csv(csv_file, index=False)


def load_report_data(file_path, date1, date2):
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
    
    df_data["filemonth_dt"] = pd.to_datetime(df_data["filemonth_dt"])
    
    summary_df = generate_summary_df(df_data, date1, date2)
    value_dist_df = generate_distribution_df(df_data, "value_dist", date1)
    pop_comp_df = generate_distribution_df(df_data, "pop_comp", date1)
    
    return df_data, summary_df, value_dist_df, pop_comp_df


def main():
    st.set_page_config(
        page_title="Official FRY14M Field Analysis Summary",
        layout="centered",
        initial_sidebar_state="expanded"
    )
    
    st.sidebar.title("File & Date Selection")
    
    folder = st.sidebar.selectbox("Select Folder", ["BDCOM", "WFHMSA"])
    
    folder_path = os.path.join(os.getcwd(), folder)
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
        
        # Summary Grid with Comments
        summary_grid = st.session_state.summary_df.copy()
        if "Comments" not in summary_grid.columns:
            summary_grid["Comments"] = ""
        summary_grid = load_comments_for_grid(summary_grid, "summary_comments.csv", ["Field Name"])
        gb_summary = GridOptionsBuilder.from_dataframe(summary_grid)
        gb_summary.configure_default_column(editable=True, singleClickEdit=True)
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
        if st.button("Save Summary Comments"):
            save_comments_for_grid(st.session_state.summary_df, "summary_comments.csv", ["Field Name"])
            st.success("Summary comments saved.")
        
        # Value Distribution Grid with Comments
        value_dist_grid = st.session_state.value_dist_df.copy()
        if "Comments" not in value_dist_grid.columns:
            value_dist_grid["Comments"] = ""
        value_dist_grid = load_comments_for_grid(value_dist_grid, "value_dist_comments.csv", ["Field Name", "Value Label"])
        gb_value = GridOptionsBuilder.from_dataframe(value_dist_grid)
        gb_value.configure_default_column(editable=True, singleClickEdit=True)
        gridOptions_value = gb_value.build()
        st.subheader("Value Distribution")
        value_response = AgGrid(
            value_dist_grid,
            gridOptions=gridOptions_value,
            update_mode=GridUpdateMode.VALUE_CHANGED,
            data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
            key="value_dist_grid"
        )
        st.session_state.value_dist_df = pd.DataFrame(value_response["data"])
        if st.button("Save Value Distribution Comments"):
            save_comments_for_grid(st.session_state.value_dist_df, "value_dist_comments.csv", ["Field Name", "Value Label"])
            st.success("Value Distribution comments saved.")
        
        # Population Comparison Grid with Comments
        pop_comp_grid = st.session_state.pop_comp_df.copy()
        if "Comments" not in pop_comp_grid.columns:
            pop_comp_grid["Comments"] = ""
        pop_comp_grid = load_comments_for_grid(pop_comp_grid, "pop_comp_comments.csv", ["Field Name", "Value Label"])
        gb_pop = GridOptionsBuilder.from_dataframe(pop_comp_grid)
        gb_pop.configure_default_column(editable=True, singleClickEdit=True)
        gridOptions_pop = gb_pop.build()
        st.subheader("Population Comparison")
        pop_response = AgGrid(
            pop_comp_grid,
            gridOptions=gridOptions_pop,
            update_mode=GridUpdateMode.VALUE_CHANGED,
            data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
            key="pop_comp_grid"
        )
        st.session_state.pop_comp_df = pd.DataFrame(pop_response["data"])
        if st.button("Save Population Comparison Comments"):
            save_comments_for_grid(st.session_state.pop_comp_df, "pop_comp_comments.csv", ["Field Name", "Value Label"])
            st.success("Population Comparison comments saved.")
        
        # Prepare Excel download using updated session_state data
        summary_updated = st.session_state.summary_df
        value_updated = st.session_state.value_dist_df
        pop_updated = st.session_state.pop_comp_df
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            summary_updated.to_excel(writer, index=False, sheet_name="Summary")
            value_updated.to_excel(writer, index=False, sheet_name="Value Distribution")
            pop_updated.to_excel(writer, index=False, sheet_name="Population Comparison")
        processed_data = output.getvalue()
        
        st.download_button(
            label="Download Report as Excel",
            data=processed_data,
            file_name="FRY14M_Field_Analysis_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


if __name__ == "__main__":
    main()