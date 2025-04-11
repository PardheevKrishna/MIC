import streamlit as st
import pandas as pd
import textwrap
from rapidfuzz import fuzz

# Set the page layout to wide.
st.set_page_config(layout="wide")

# Helper function to insert newline characters every 'width' characters.
def wrap_text(text, width=100):
    # This inserts newlines every 'width' characters.
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
        # Iterate over each row in the DataFrame.
        for idx, row in df.iterrows():
            row_matched = False
            # For each cell, split the cell text into words and compare.
            for col in df.columns:
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
                row_dict = row.to_dict()
                row_dict["Row Index"] = idx
                row_dict["Source"] = source_name
                row_dict["Sheet"] = sheet_name  # in case you need it in display
                sheet_results.append(row_dict)
        if sheet_results:
            results[sheet_name] = results.get(sheet_name, []) + sheet_results
    return results

def main():
    st.title("Excel Search Application")
    st.write("Upload Excel files and enter a search term to search every word in every cell for fuzzy matches.")

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
            # Get the unique source files for this sheet.
            sources = sorted({row["Source"] for row in rows})
            expander_title = f"Sheet: {sheet_name} ({len(rows)} match{'es' if len(rows) > 1 else ''})"
            expander_title += " - Files: " + ", ".join(sources)
            with st.expander(expander_title, expanded=True):
                # Before creating a DataFrame, drop unnamed columns that are completely empty.
                # We'll define a helper to do that for DataFrames.
                def drop_empty_unnamed_cols(df):
                    # Drop any column whose name starts with "Unnamed" and has all NaN values.
                    unnamed = [col for col in df.columns if col.startswith("Unnamed")]
                    for col in unnamed:
                        if df[col].isnull().all():
                            df = df.drop(columns=[col])
                    return df

                if sheet_name == "1. Data Dictionary":
                    # For this sheet, display only the required columns.
                    required_cols = [
                        "CorporateFinanceSubmissionFieldName",
                        "Corporate Finance Submission Field Description",
                        "Transformation/Business Logic"
                    ]
                    # Build display headers. The second header is modified to have a newline.
                    display_headers = ["Source", "CorporateFinanceSubmissionFieldName", 
                                       "Corporate Finance Submission\nField Description", 
                                       "Transformation/Business Logic"]
                    display_rows = []
                    for row in rows:
                        # Create a display row with only required columns.
                        display_row = {
                            "Source": row.get("Source", ""),
                            "CorporateFinanceSubmissionFieldName": row.get("CorporateFinanceSubmissionFieldName", ""),
                            "Corporate Finance Submission Field Description": row.get("Corporate Finance Submission Field Description", ""),
                            "Transformation/Business Logic": row.get("Transformation/Business Logic", "")
                        }
                        # For each cell, if text is longer than 100 characters, insert a newline every 100 characters.
                        for key, val in display_row.items():
                            if isinstance(val, str) and len(val) > 100:
                                display_row[key] = wrap_text(val, width=100)
                        display_rows.append(display_row)
                    df_display = pd.DataFrame(display_rows, columns=display_headers)
                    df_display = drop_empty_unnamed_cols(df_display)
                    # Use pandas Styler to set the CSS for word wrapping.
                    df_display_styled = df_display.style.set_properties(**{'white-space': 'normal'})
                    st.dataframe(df_display_styled, use_container_width=True)
                else:
                    # For other sheets, display all columns (including Source and Sheet).
                    df_display = pd.DataFrame(rows)
                    df_display = drop_empty_unnamed_cols(df_display)
                    df_display_styled = df_display.style.set_properties(**{'white-space': 'normal'})
                    st.dataframe(df_display_styled, use_container_width=True)

if __name__ == "__main__":
    main()