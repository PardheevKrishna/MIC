import streamlit as st
import pandas as pd
import textwrap
from rapidfuzz import fuzz

# Set the layout to wide.
st.set_page_config(layout="wide")

# Helper function that inserts newline characters every 'width' characters.
def wrap_text(text, width=100):
    return "\n".join(text[i:i+width] for i in range(0, len(text), width))

# Process a single Excel file.
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
        # Drop columns that are entirely empty.
        df_clean = df.dropna(axis=1, how='all')
        sheet_results = []
        # Iterate over each row.
        for idx, row in df_clean.iterrows():
            row_matched = False
            # For each cell in the row, split the cell text into words and compare.
            for col in df_clean.columns:
                cell_value = row[col]
                if pd.isna(cell_value):
                    continue
                cell_text = str(cell_value)
                words = cell_text.split()
                for word in words:
                    if fuzz.ratio(search_term.lower(), word.lower()) >= 80:
                        row_matched = True
                        break
                if row_matched:
                    break
            if row_matched:
                # Convert the row to a dictionary.
                row_dict = row_clean = row.to_dict()
                row_dict["Row Index"] = idx
                row_dict["Source"] = source_name
                row_dict["Sheet"] = sheet_name
                sheet_results.append(row_dict)
        if sheet_results:
            # Preserve all non-empty columns.
            results[sheet_name] = results.get(sheet_name, []) + sheet_results
    return results

def main():
    st.title("Excel Search Application")
    st.write("Upload Excel files and enter a search term. The app searches every word in every cell for fuzzy matches, drops completely empty columns, and adjusts cell height so that long text wraps automatically.")

    # File uploader that accepts multiple Excel files.
    uploaded_files = st.file_uploader("Upload Excel files", type=["xlsx", "xls"], accept_multiple_files=True)
    search_term = st.text_input("Enter search term")
    
    if st.button("Search"):
        if not uploaded_files:
            st.warning("Please upload at least one Excel file.")
            return
        if not search_term:
            st.warning("Please enter a search term.")
            return

        aggregated_results = {}  # Dictionary: {sheet_name: list of row dicts}
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
        # Display results grouped by sheet.
        for sheet_name, rows in aggregated_results.items():
            # Create an expander for each sheet. Also show all the sources for that sheet.
            sources = sorted({row["Source"] for row in rows})
            expander_title = f"Sheet: {sheet_name} ({len(rows)} match{'es' if len(rows) > 1 else ''}) - Files: " + ", ".join(sources)
            with st.expander(expander_title, expanded=True):
                if sheet_name == "1. Data Dictionary":
                    # For "1. Data Dictionary", keep only required columns.
                    required_cols = [
                        "CorporateFinanceSubmissionFieldName",
                        "Corporate Finance Submission Field Description",
                        "Transformation/Business Logic"
                    ]
                    # Define display headers with an explicit newline in the second header.
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
                        # Wrap text for each cell value longer than 100 characters.
                        for key, val in display_row.items():
                            if isinstance(val, str) and len(val) > 100:
                                display_row[key] = wrap_text(val, width=100)
                        display_rows.append(display_row)
                    df_display = pd.DataFrame(display_rows, columns=display_headers)
                    # Use a large height so that row heights adjust to content.
                    st.dataframe(df_display, use_container_width=True, height=600)
                else:
                    # For all other sheets, display all columns (drop empty ones).
                    df_display = pd.DataFrame(rows)
                    df_display = df_display.dropna(axis=1, how='all')
                    st.dataframe(df_display, use_container_width=True, height=600)

if __name__ == "__main__":
    main()