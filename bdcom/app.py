import streamlit as st
st.set_page_config(page_title="Official FRY14M Field Analysis Summary", layout="centered", initial_sidebar_state="expanded")

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

        summary_data.append([field, missing_sum_d1, missing_sum_d2, m2m_sum_d1, m2m_sum_d2])
    return pd.DataFrame(
        summary_data,
        columns=[
            "Field Name",
            f"Missing Values ({date1.strftime('%Y-%m-%d')})",
            f"Missing Values ({date2.strftime('%Y-%m-%d')})",
            f"Month-to-Month Diff ({date1.strftime('%Y-%m-%d')})",
            f"Month-to-Month Diff ({date2.strftime('%Y-%m-%d')})"
        ]
    )

def generate_distribution_df(df, analysis_type, date1):
    months = [(date1 - relativedelta(months=i)).replace(day=1) for i in range(12)]
    months = sorted(months, reverse=True)
    df_filtered = df[df['analysis_type'] == analysis_type].copy()
    df_filtered['month'] = df_filtered['filemonth_dt'].apply(lambda d: d.replace(day=1))
    df_filtered = df_filtered[df_filtered['month'].isin(months)]
    grouped = df_filtered.groupby(['field_name', 'value_label', 'month'])['value_records'].sum().reset_index()
    if grouped.empty:
        return pd.DataFrame()

    pivot = grouped.pivot_table(
        index=['field_name', 'value_label'],
        columns='month',
        values='value_records',
        fill_value=0
    )
    pivot = pivot.reindex(columns=months, fill_value=0)
    result_frames = []
    for field, sub_df in pivot.groupby(level=0):
        sub_df = sub_df.droplevel(0)
        total = sub_df.sum(axis=0)
        percent_df = sub_df.div(total, axis=1).mul(100).round(2).fillna(0)
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
        temp_df.index = pd.MultiIndex.from_product(
            [[field], temp_df.index],
            names=["Field Name", "Value Label"]
        )
        result_frames.append(temp_df)
    if not result_frames:
        return pd.DataFrame()

    final_df = pd.concat(result_frames)
    final_df.columns = pd.MultiIndex.from_tuples(final_df.columns)
    return final_df

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
    value_dist_df = generate_distribution_df(df_data, "value_dist", date1)
    pop_comp_df = generate_distribution_df(df_data, "pop_comp", date1)
    return df_data, summary_df, value_dist_df, pop_comp_df

def main():
    st.sidebar.title("File & Date Selection")
    folder = st.sidebar.selectbox("Select Folder", ["BDCOM", "WFHMSA"])
    folder_path = os.path.join(os.getcwd(), folder)
    if not os.path.exists(folder_path):
        st.sidebar.error(f"Folder '{folder}' not found.")
        return
    all_files = [f for f in os.listdir(folder_path) if f.lower().endswith(('.xlsx', '.xlsb'))]
    if not all_files:
        st.sidebar.error(f"No Excel files found in '{folder}' folder.")
        return

    selected_file = st.sidebar.selectbox("Select an Excel File", all_files)
    selected_date = st.sidebar.date_input("Select Date for Date1", datetime.date(2025, 1, 1))
    date1 = datetime.datetime.combine(selected_date, datetime.datetime.min.time())
    date2 = date1 - relativedelta(months=1)

    if st.sidebar.button("Generate Report"):
        df_data, summary_df, value_dist_df, pop_comp_df = load_report_data(
            os.path.join(folder_path, selected_file), date1, date2
        )
        st.session_state.report_generated = True
        st.session_state.df_data = df_data
        st.session_state.summary_df = summary_df
        st.session_state.value_dist_df = flatten_dataframe(value_dist_df.copy())
        st.session_state.pop_comp_df = flatten_dataframe(pop_comp_df.copy())
        st.session_state.selected_file = selected_file
        st.session_state.folder = folder
        st.session_state.date1 = date1
        st.session_state.date2 = date2
        st.session_state.current_field_value = 0
        st.session_state.current_field_pop = 0

    if st.session_state.get("report_generated", False):
        st.title("Official FRY14M Field Analysis Summary Report")
        st.write(f"**Folder:** {st.session_state.folder}")
        st.write(f"**Selected File:** {st.session_state.selected_file}")
        st.write(f"**Date1:** {st.session_state.date1.strftime('%Y-%m-%d')} | **Date2:** {st.session_state.date2.strftime('%Y-%m-%d')}")

        summary_df = st.session_state.summary_df.copy()
        if "Comments" not in summary_df.columns:
            summary_df["Comments"] = ""
        gb_sum = GridOptionsBuilder.from_dataframe(summary_df)
        gb_sum.configure_default_column(editable=False)
        gb_sum.configure_column("Comments", editable=True)
        gb_sum.configure_selection('single', use_checkbox=False)
        gridOptions_sum = gb_sum.build()
        gridOptions_sum["rowSelection"] = "single"

        st.subheader("Summary")
        sum_response = AgGrid(
            summary_df,
            gridOptions=gridOptions_sum,
            update_mode=GridUpdateMode.SELECTION_CHANGED,
            data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
            key="summary_grid"
        )
        sum_selected = sum_response.get("selectedRows", [])
        if sum_selected:
            field_link = sum_selected[0].get("Field Name")
            val_fields = st.session_state.value_dist_df["Field Name"].unique().tolist()
            pop_fields = st.session_state.pop_comp_df["Field Name"].unique().tolist()
            if field_link in val_fields:
                st.session_state.current_field_value = val_fields.index(field_link)
            if field_link in pop_fields:
                st.session_state.current_field_pop = pop_fields.index(field_link)

        st.subheader("Value Distribution")
        val_fields = st.session_state.value_dist_df["Field Name"].unique().tolist()
        col1, col2 = st.columns(2)
        if col1.button("← Value", key="prev_value"):
            st.session_state.current_field_value = (st.session_state.current_field_value - 1) % len(val_fields)
        if col2.button("Value →", key="next_value"):
            st.session_state.current_field_value = (st.session_state.current_field_value + 1) % len(val_fields)

        selected_field_val = st.selectbox(
            "Select Field (Value Dist)",
            val_fields,
            index=st.session_state.current_field_value,
            key="val_field_select"
        )
        st.session_state.current_field_value = val_fields.index(selected_field_val)

        filtered_val_dist = st.session_state.value_dist_df[
            st.session_state.value_dist_df["Field Name"] == selected_field_val
        ].copy()
        if "Comments" not in filtered_val_dist.columns:
            filtered_val_dist["Comments"] = ""
        gb_val = GridOptionsBuilder.from_dataframe(filtered_val_dist)
        gb_val.configure_default_column(editable=False)
        gb_val.configure_column("Comments", editable=True)
        gb_val.configure_selection('single', use_checkbox=False)
        gridOptions_val = gb_val.build()
        gridOptions_val["rowSelection"] = "single"

        val_response = AgGrid(
            filtered_val_dist,
            gridOptions=gridOptions_val,
            update_mode=GridUpdateMode.SELECTION_CHANGED,
            data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
            key="val_dist_grid"
        )
        val_selected = val_response.get("selectedRows", [])
        if val_selected:
            field_sel = val_selected[0].get("Field Name")
            pop_fields = st.session_state.pop_comp_df["Field Name"].unique().tolist()
            if field_sel in pop_fields:
                st.session_state.current_field_pop = pop_fields.index(field_sel)

        st.subheader("Population Comparison")
        pop_fields = st.session_state.pop_comp_df["Field Name"].unique().tolist()
        col3, col4 = st.columns(2)
        if col3.button("← Pop", key="prev_pop"):
            st.session_state.current_field_pop = (st.session_state.current_field_pop - 1) % len(pop_fields)
        if col4.button("Pop →", key="next_pop"):
            st.session_state.current_field_pop = (st.session_state.current_field_pop + 1) % len(pop_fields)

        selected_field_pop = st.selectbox(
            "Select Field (Pop Comp)",
            pop_fields,
            index=st.session_state.current_field_pop,
            key="pop_field_select"
        )
        st.session_state.current_field_pop = pop_fields.index(selected_field_pop)

        filtered_pop = st.session_state.pop_comp_df[
            st.session_state.pop_comp_df["Field Name"] == selected_field_pop
        ].copy()
        if "Comments" not in filtered_pop.columns:
            filtered_pop["Comments"] = ""

        gb_pop = GridOptionsBuilder.from_dataframe(filtered_pop)
        gb_pop.configure_default_column(editable=False)
        gb_pop.configure_column("Comments", editable=True)
        # No "Show SQL" checkbox; we'll show logic automatically on row selection
        gb_pop.configure_selection('single', use_checkbox=False)
        gridOptions_pop = gb_pop.build()
        gridOptions_pop["rowSelection"] = "single"

        pop_response = AgGrid(
            filtered_pop,
            gridOptions=gridOptions_pop,
            update_mode=GridUpdateMode.SELECTION_CHANGED,
            data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
            key="pop_comp_grid"
        )
        pop_selected = pop_response.get("selectedRows", [])
        if pop_selected:
            field_sel = pop_selected[0].get("Field Name")
            value_label = pop_selected[0].get("Value Label")
            # Show SQL automatically if not "Current period total"
            if value_label != "Current period total":
                df_data = st.session_state.df_data
                logic_rows = df_data[
                    (df_data["analysis_type"] == "pop_comp") &
                    (df_data["field_name"] == field_sel) &
                    (df_data["value_label"] == value_label)
                ]
                sql_vals = logic_rows["value_sql_logic"].unique()
                if sql_vals.size > 0:
                    st.text_area("Value SQL Logic", "\n".join(sql_vals), height=100)
                else:
                    st.text_area("Value SQL Logic", "No SQL Logic found", height=100)

        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            st.session_state.summary_df.to_excel(writer, index=False, sheet_name="Summary")
            st.session_state.value_dist_df.to_excel(writer, index=False, sheet_name="Value Distribution")
            st.session_state.pop_comp_df.to_excel(writer, index=False, sheet_name="Population Comparison")
        st.download_button(
            "Download Report as Excel",
            data=output.getvalue(),
            file_name="FRY14M_Field_Analysis_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()