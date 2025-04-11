import streamlit as st
import pandas as pd
from rapidfuzz import fuzz
import textwrap
import streamlit.components.v1 as components

# Set the page configuration to wide.
st.set_page_config(layout="wide", page_title="Excel Search Application")

##########################################
# Helper function: Wrap text into HTML.
##########################################
def wrap_text_html(text, width=100):
    # Insert <br> tags every 'width' characters.
    return "<br>".join(text[i:i+width] for i in range(0, len(text), width))

##########################################
# Helper function: Generate HTML table from a DataFrame.
##########################################
def generate_html_table_from_df(df, wrap_width=100):
    def wrap_cell(val):
        s = str(val)
        if len(s) > wrap_width:
            return "<br>".join(s[i:i+wrap_width] for i in range(0, len(s), wrap_width))
        else:
            return s

    # Inline CSS to force cell content to wrap and adjust the row height.
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
    <table>
      <thead>
        <tr>
    """
    for col in df.columns:
        html += f"<th>{col}</th>"
    html += "</tr></thead><tbody>"
    for i, row in df.iterrows():
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

    for sheet_name, df in excel_data.items():
        # Drop columns that are entirely empty.
        df_clean = df.dropna(axis=1, how='all')
        sheet_results = []
        for idx, row in df_clean.iterrows():
            row_matched = False
            # For each cell, check every word.
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
            results[sheet_name] = results.get(sheet_name, []) + sheet_results
    return results

##########################################
# Main function.
##########################################
def main():
    st.title("Excel Search Application")
    st.write(
        "Upload Excel files and enter a search term. "
        "The app searches every word in every cell (using fuzzy matching) for a match. "
        "Empty columns are dropped. All matching rows include the Source file name and Sheet name. "
        "Cell content is wrapped (with a line break inserted every 100 characters) so that each cell's height adjusts automatically."
    )

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

        aggregated_results = {}  # {sheet_name: list of row dictionaries}
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
                if sheet_name == "1. Data Dictionary":
                    # For Data Dictionary, only show the required columns.
                    required_cols = [
                        "CorporateFinanceSubmissionFieldName",
                        "Corporate Finance Submission Field Description",
                        "Transformation/Business Logic"
                    ]
                    # Modify header for the second required column.
                    display_headers = ["Source", 
                                       "CorporateFinanceSubmissionFieldName", 
                                       "Corporate Finance Submission<br>Field Description", 
                                       "Transformation/Business Logic"]
                    display_rows = []
                    for row in rows:
                        # Only include rows where at least one required field is non-empty.
                        if any(str(row.get(col, "")).strip() for col in required_cols):
                            display_row = {
                                "Source": row.get("Source", ""),
                                "CorporateFinanceSubmissionFieldName": row.get("CorporateFinanceSubmissionFieldName", ""),
                                "Corporate Finance Submission Field Description": row.get("Corporate Finance Submission Field Description", ""),
                                "Transformation/Business Logic": row.get("Transformation/Business Logic", "")
                            }
                            display_rows.append(display_row)
                    if display_rows:
                        df_display = pd.DataFrame(display_rows, columns=display_headers)
                        html_table = generate_html_table_from_df(df_display, wrap_width=100)
                        # Render the HTML table using components.html
                        components.html(html_table, height=700, scrolling=True)
                    else:
                        st.info("No non-empty rows found for required columns.")
                else:
                    # For other sheets, display all non-empty columns.
                    df_display = pd.DataFrame(rows)
                    df_display = df_display.dropna(axis=1, how='all')
                    html_table = generate_html_table_from_df(df_display, wrap_width=100)
                    components.html(html_table, height=700, scrolling=True)

if __name__ == "__main__":
    main()