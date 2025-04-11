import streamlit as st
import pandas as pd
from rapidfuzz import fuzz

# Set the page layout to wide.
st.set_page_config(layout="wide", page_title="Excel Search Application")

##########################################
# Helper function to wrap text using HTML.
##########################################
def wrap_text_html(text, width=100):
    # Insert <br> tags every 'width' characters.
    return "<br>".join(text[i:i+width] for i in range(0, len(text), width))

##########################################
# Generate an HTML table from a DataFrame.
##########################################
def generate_html_table_from_df(df, wrap_width=100):
    # Apply word wrapping to each cell if needed.
    def wrap_cell(val):
        s = str(val)
        if len(s) > wrap_width:
            return wrap_text_html(s, width=wrap_width)
        else:
            return s

    # Define CSS to ensure cells wrap and height adjusts.
    html = """
    <style>
      table {
        width: 100%;
        border-collapse: collapse;
        table-layout: auto;
      }
      th, td {
        border: 1px solid #ccc;
        padding: 8px;
        white-space: pre-wrap;
        word-wrap: break-word;
        vertical-align: top;
      }
      th {
        background-color: #3c3c4e;
        color: #ffffff;
      }
    </style>
    """
    html += "<table><thead><tr>"
    # Create table header.
    for col in df.columns:
        html += f"<th>{col}</th>"
    html += "</tr></thead><tbody>"
    # Create table rows.
    for index, row in df.iterrows():
        html += "<tr>"
        for col in df.columns:
            cell_val = wrap_cell(row[col])
            html += f"<td>{cell_val}</td>"
        html += "</tr>"
    html += "</tbody></table>"
    return html

##########################################
# Process a single Excel file.
##########################################
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
        # Drop columns that are completely empty.
        df_clean = df.dropna(axis=1, how='all')
        sheet_results = []
        # Process each row.
        for idx, row in df_clean.iterrows():
            row_matched = False
            # Check each cell: split into words and compare each word.
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
                row_dict = row.to_dict()
                row_dict["Row Index"] = idx
                row_dict["Source"] = source_name
                row_dict["Sheet"] = sheet_name
                sheet_results.append(row_dict)
        if sheet_results:
            # Append results for the sheet.
            results[sheet_name] = results.get(sheet_name, []) + sheet_results
    return results

##########################################
# Main function
##########################################
def main():
    st.title("Excel Search Application")
    st.write("Upload Excel files and enter a search term. The search is performed word-by-word in every cell. "
             "Empty columns are dropped. The results for each sheet display the Source file and Sheet name. "
             "Cell content is wrapped (newlines inserted every 100 characters) so that row height adjusts automatically.")

    uploaded_files = st.file_uploader("Upload Excel files", type=["xlsx", "xls"], accept_multiple_files=True)
    search_term = st.text_input("Enter search term")

    if st.button("Search"):
        if not uploaded_files:
            st.warning("Please upload at least one Excel file.")
            return
        if not search_term:
            st.warning("Please enter a search term.")
            return

        aggregated_results = {}  # {sheet_name: list of row dicts}
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
            sources = sorted({row["Source"] for row in rows})
            expander_title = f"Sheet: {sheet_name} ({len(rows)} match{'es' if len(rows) > 1 else ''}) - Files: " + ", ".join(sources)
            with st.expander(expander_title, expanded=True):
                # For the "1. Data Dictionary" sheet, display only required columns.
                if sheet_name == "1. Data Dictionary":
                    required_cols = [
                        "CorporateFinanceSubmissionFieldName",
                        "Corporate Finance Submission Field Description",
                        "Transformation/Business Logic"
                    ]
                    # Build a DataFrame with only the required columns plus Source.
                    display_rows = []
                    for row in rows:
                        # Check if at least one required column is non-empty.
                        if any(str(row.get(col, "")).strip() for col in required_cols):
                            display_row = {
                                "Source": row.get("Source", ""),
                                "CorporateFinanceSubmissionFieldName": row.get("CorporateFinanceSubmissionFieldName", ""),
                                "Corporate Finance Submission Field Description": row.get("Corporate Finance Submission Field Description", ""),
                                "Transformation/Business Logic": row.get("Transformation/Business Logic", "")
                            }
                            display_rows.append(display_row)
                    if display_rows:
                        df_display = pd.DataFrame(display_rows)
                        # Rename header for the second column to include a line break.
                        df_display = df_display.rename(
                            columns={"Corporate Finance Submission Field Description":
                                     "Corporate Finance Submission<br>Field Description"})
                        # Generate HTML table with wrapped content.
                        html_table = generate_html_table_from_df(df_display, wrap_width=100)
                        st.markdown(html_table, unsafe_allow_html=True)
                    else:
                        st.info("No non-empty rows found for required columns.")
                else:
                    # For other sheets, display all non-empty columns.
                    df_display = pd.DataFrame(rows)
                    df_display = df_display.dropna(axis=1, how='all')
                    html_table = generate_html_table_from_df(df_display, wrap_width=100)
                    st.markdown(html_table, unsafe_allow_html=True)

if __name__ == "__main__":
    main()