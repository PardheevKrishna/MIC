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
        sheet_results = []
        # Iterate over each row.
        for idx, row in df.iterrows():
            row_matched = False
            # For each cell, split the cell text into words and compare.
            for col in df.columns:
                cell_value = row[col]
                if pd.isna(cell_value):
                    continue
                cell_text = str(cell_value)
                words = cell_text.split()
                # If any word in the cell is similar enough, mark the row as matching.
                for word in words:
                    if fuzz.ratio(search_term.lower(), word.lower()) >= 80:
                        row_matched = True
                        break
                if row_matched:
                    break
            if row_matched:
                # Convert the row to a dictionary.
                row_dict = row.to_dict()
                row_dict["Row Index"] = idx
                row_dict["Source"] = source_name
                row_dict["Sheet"] = sheet_name  # Include sheet name for display.
                sheet_results.append(row_dict)
        if sheet_results:
            results[sheet_name] = results.get(sheet_name, []) + sheet_results
    return results

def main():
    st.title("Excel Search Application")
    st.write("Upload Excel files and enter a search term to search every word in every cell for fuzzy matches.")
    
    # File uploader that accepts multiple files.
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
            # Create an expander for each sheet. Show sheet name plus list of sources.
            sources = sorted({row["Source"] for row in rows})
            expander_title = f"Sheet: {sheet_name} ({len(rows)} match{'es' if len(rows) > 1 else ''})"
            expander_title += " - Files: " + ", ".join(sources)
            with st.expander(expander_title, expanded=True):
                if sheet_name == "1. Data Dictionary":
                    # For the "1. Data Dictionary" sheet, display only the required columns.
                    required_cols = [
                        "CorporateFinanceSubmissionFieldName",
                        "Corporate Finance Submission Field Description",
                        "Transformation/Business Logic"
                    ]
                    # Build display headers. Insert a newline in the second header.
                    display_headers = ["Source"] + [
                        "CorporateFinanceSubmissionFieldName", 
                        "Corporate Finance Submission\nField Description", 
                        "Transformation/Business Logic"
                    ]
                    display_rows = []
                    for row in rows:
                        display_row = {
                            "Source": row.get("Source", ""),
                            "CorporateFinanceSubmissionFieldName": row.get("CorporateFinanceSubmissionFieldName", ""),
                            "Corporate Finance Submission Field Description": row.get("Corporate Finance Submission Field Description", ""),
                            "Transformation/Business Logic": row.get("Transformation/Business Logic", "")
                        }
                        # Wrap text for each cell value that exceeds 100 characters.
                        for key, val in display_row.items():
                            if isinstance(val, str) and len(val) > 100:
                                display_row[key] = wrap_text(val, width=100)
                        display_rows.append(display_row)
                    df_display = pd.DataFrame(display_rows, columns=display_headers)
                    st.dataframe(df_display, use_container_width=True)
                else:
                    # For all other sheets, display all columns.
                    # Ensure "Source" and "Sheet" are included.
                    for row in rows:
                        if "Source" not in row:
                            row["Source"] = ""
                        if "Sheet" not in row:
                            row["Sheet"] = sheet_name
                    df_display = pd.DataFrame(rows)
                    st.dataframe(df_display, use_container_width=True)

if __name__ == "__main__":
    main()