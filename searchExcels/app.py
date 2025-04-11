import streamlit as st
import pandas as pd
from rapidfuzz import fuzz

# Set the page layout to wide.
st.set_page_config(layout="wide", page_title="Excel Search Application")

# Helper function: Insert HTML <br> tags every 'width' characters.
def wrap_text_html(text, width=100):
    return "<br>".join(text[i:i+width] for i in range(0, len(text), width))

# Process a single Excel file.
# Reads all sheets, drops columns that are entirely empty, and checks each cell's words for a fuzzy match.
def process_excel_file(uploaded_file, source_name, search_term):
    results = {}
    try:
        # Read all sheets from the Excel file.
        excel_data = pd.read_excel(uploaded_file, sheet_name=None)
    except Exception as e:
        st.error(f"Error reading {source_name}: {e}")
        return results

    for sheet_name, df in excel_data.items():
        # Drop columns that are completely empty.
        df_clean = df.dropna(axis=1, how='all')
        sheet_results = []
        # Process each row.
        for idx, row in df_clean.iterrows():
            row_matched = False
            # For each cell, split into words and compare each word.
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

# Generate an HTML table for the "1. Data Dictionary" sheet.
# Only required columns are shown, empty rows (all required fields empty) are dropped,
# and each cell's text is wrapped (with <br>) dynamically every 100 characters.
def generate_data_dictionary_html(rows, required_cols, header_mapping, wrap_width=100):
    # Filter out rows where all required columns are empty.
    filtered_rows = []
    for row in rows:
        non_empty = any(str(row.get(col, "")).strip() for col in required_cols)
        if non_empty:
            filtered_rows.append(row)
    if not filtered_rows:
        return "<p>No matches found.</p>"
    
    # Build HTML table with inline CSS.
    html = """
    <style>
      table {
        width: 100%;
        border-collapse: collapse;
        table-layout: fixed;
      }
      th, td {
        border: 1px solid #ccc;
        padding: 8px;
        white-space: pre-wrap;
        vertical-align: top;
        word-wrap: break-word;
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
    # Table headers.
    headers = ["Source"] + [header_mapping.get(col, col) for col in required_cols]
    for h in headers:
        html += f"<th>{h}</th>"
    html += "</tr></thead><tbody>"
    
    # Table rows.
    for row in filtered_rows:
        html += "<tr>"
        html += f"<td>{row.get('Source', '')}</td>"
        for col in required_cols:
            cell_val = str(row.get(col, ""))
            if len(cell_val) > wrap_width:
                cell_val = wrap_text_html(cell_val, width=wrap_width)
            html += f"<td>{cell_val}</td>"
        html += "</tr>"
    html += "</tbody></table>"
    return html

def main():
    st.title("Excel Search Application")
    st.write("Upload Excel files and enter a search term. The app searches every word in every cell for fuzzy matches. "
             "Empty columns are dropped. For the '1. Data Dictionary' sheet, only specific columns are displayed, and "
             "cell text is wrapped (newline inserted every 100 characters) so rows expand vertically.")

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
                if sheet_name == "1. Data Dictionary":
                    required_cols = [
                        "CorporateFinanceSubmissionFieldName",
                        "Corporate Finance Submission Field Description",
                        "Transformation/Business Logic"
                    ]
                    # Map header for "Corporate Finance Submission Field Description" to include a line break.
                    header_mapping = {
                        "CorporateFinanceSubmissionFieldName": "CorporateFinanceSubmissionFieldName",
                        "Corporate Finance Submission Field Description": "Corporate Finance Submission<br>Field Description",
                        "Transformation/Business Logic": "Transformation/Business Logic"
                    }
                    html = generate_data_dictionary_html(rows, required_cols, header_mapping, wrap_width=100)
                    st.markdown(html, unsafe_allow_html=True)
                else:
                    # For other sheets, drop columns that are completely empty.
                    df_display = pd.DataFrame(rows)
                    df_display = df_display.dropna(axis=1, how='all')
                    st.dataframe(df_display, use_container_width=True)

if __name__ == "__main__":
    main()