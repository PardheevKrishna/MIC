import streamlit as st
import pandas as pd
import textwrap
from rapidfuzz import fuzz

# Set the page to use a wide layout.
st.set_page_config(layout="wide")

# Helper function to insert newline characters every 'width' characters.
def wrap_text(text, width=100):
    return "\n".join(text[i:i+width] for i in range(0, len(text), width))

def process_excel_file(uploaded_file, source_name, search_term):
    results = {}
    try:
        # Read all sheets from the Excel file.
        excel_data = pd.read_excel(uploaded_file, sheet_name=None)
    except Exception as e:
        st.error(f"Error reading {source_name}: {e}")
        return results

    # Loop over each sheet.
    for sheet_name, df in excel_data.items():
        sheet_results = []
        for idx, row in df.iterrows():
            # Filter out any columns whose key starts with "Unnamed:".
            row_dict = {k: v for k, v in row.to_dict().items() if not str(k).startswith("Unnamed:")}
            row_matched = False
            # For each cell in the row, split the cell text into words and compare each word.
            for col in df.columns:
                if col.startswith("Unnamed:"):
                    continue
                cell_value = row.get(col)
                if pd.isna(cell_value):
                    continue
                cell_str = str(cell_value)
                words = cell_str.split()
                for word in words:
                    if fuzz.ratio(search_term.lower(), word.lower()) >= 80:
                        row_matched = True
                        break
                if row_matched:
                    break
            if row_matched:
                row_dict["Row Index"] = idx
                row_dict["Source"] = source_name
                sheet_results.append(row_dict)
        if sheet_results:
            results[sheet_name] = results.get(sheet_name, []) + sheet_results
    return results

def main():
    st.title("Excel Search Application")
    st.write("Upload Excel files and enter a search term to search every cell (word‐by‐word) for fuzzy matches.")

    uploaded_files = st.file_uploader("Upload Excel files", type=["xlsx", "xls"], accept_multiple_files=True)
    search_term = st.text_input("Enter search term")
    if st.button("Search"):
        if not uploaded_files:
            st.warning("Please upload at least one Excel file.")
            return
        if not search_term:
            st.warning("Please enter a search term.")
            return

        aggregated_results = {}  # Dictionary: {sheet_name: list of row dictionaries}
        for uploaded_file in uploaded_files:
            source_name = uploaded_file.name
            st.write(f"Processing file: {source_name}")
            file_results = process_excel_file(uploaded_file, source_name, search_term)
            for sheet_name, rows in file_results.items():
                if sheet_name not in aggregated_results:
                    aggregated_results[sheet_name] = []
                aggregated_results[sheet_name].extend(rows)
        if not aggregated_results:
            st.info("No matches found.")
            return

        st.write("### Search Results")
        # Display the results for each sheet in an expander.
        for sheet_name, rows in aggregated_results.items():
            with st.expander(f"Sheet: {sheet_name} ({len(rows)} match{'es' if len(rows) > 1 else ''})", expanded=True):
                if sheet_name == "1. Data Dictionary":
                    # For the "1. Data Dictionary" sheet, display only the required columns.
                    required_cols = [
                        "CorporateFinanceSubmissionFieldName",
                        "Corporate Finance Submission Field Description",
                        "Transformation/Business Logic"
                    ]
                    # Set header: the second column header is modified to include a newline.
                    display_headers = ["Source", 
                                       "CorporateFinanceSubmissionFieldName", 
                                       "Corporate Finance Submission\nField Description", 
                                       "Transformation/Business Logic"]
                    display_rows = []
                    for row in rows:
                        display_row = {
                            "Source": row.get("Source", ""),
                            "CorporateFinanceSubmissionFieldName": row.get("CorporateFinanceSubmissionFieldName", ""),
                            "Corporate Finance Submission Field Description": row.get("Corporate Finance Submission Field Description", ""),
                            "Transformation/Business Logic": row.get("Transformation/Business Logic", "")
                        }
                        # For each cell, if text exceeds 100 characters, insert newline breaks every 100 characters.
                        for key, val in display_row.items():
                            if isinstance(val, str) and len(val) > 100:
                                display_row[key] = wrap_text(val, width=100)
                        display_rows.append(display_row)
                    df_display = pd.DataFrame(display_rows, columns=display_headers)
                    st.dataframe(df_display, use_container_width=True)
                else:
                    # For all other sheets, ensure the "Source" column appears.
                    df_display = pd.DataFrame(rows)
                    if "Source" in df_display.columns:
                        cols = list(df_display.columns)
                        cols.remove("Source")
                        df_display = df_display[["Source"] + cols]
                    st.dataframe(df_display, use_container_width=True)

if __name__ == "__main__":
    main()