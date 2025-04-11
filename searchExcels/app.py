import streamlit as st
import pandas as pd
from rapidfuzz import fuzz
import textwrap
import streamlit.components.v1 as components

# Set the page configuration to wide.
st.set_page_config(layout="wide", page_title="Excel Search Application")

###############################################
# Helper function: Wrap text using HTML <br>.
###############################################
def wrap_text_html(text, width=100):
    return "<br>".join(text[i:i+width] for i in range(0, len(text), width))

#####################################################
# Helper function: Generate an HTML table from a DataFrame.
#####################################################
def generate_html_table_from_df(df, wrap_width=100):
    def wrap_cell(val):
        s = str(val)
        if len(s) > wrap_width:
            return "<br>".join(s[i:i+wrap_width] for i in range(0, len(s), wrap_width))
        else:
            return s

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
    for _, row in df.iterrows():
        html += "<tr>"
        for col in df.columns:
            cell_val = wrap_cell(row[col])
            html += f"<td>{cell_val}</td>"
        html += "</tr>"
    html += "</tbody></table>"
    return html

#####################################################
# Process a single Excel file.
#####################################################
def process_excel_file(uploaded_file, source_name, search_term):
    results = {}
    try:
        # Read all sheets.
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
            # Check each cell by splitting into words.
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
                # Convert the row to a dictionary with stripped keys.
                row_dict = {str(k).strip(): v for k, v in row.to_dict().items()}
                row_dict["Row Index"] = idx
                row_dict["Source"] = source_name
                row_dict["Sheet"] = sheet_name
                sheet_results.append(row_dict)
        if sheet_results:
            results[sheet_name] = results.get(sheet_name, []) + sheet_results
    return results

#####################################################
# Main function.
#####################################################
def main():
    st.title("Excel Search Application")
    st.write(
        "Upload Excel files and enter a search term. The app searches every word in every cell "
        "for fuzzy matches, drops completely empty columns, and displays the Source (file name) and Sheet name. "
        "The display for any sheet uses an HTML table where cell content is wrapped by inserting <br> every 100 characters, "
        "so row heights adjust automatically."
    )

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
        # For each sheet, display results.
        for sheet_name, rows in aggregated_results.items():
            sources = sorted({row["Source"] for row in rows})
            expander_title = f"Sheet: {sheet_name} ({len(rows)} match{'es' if len(rows) > 1 else ''}) - Files: " + ", ".join(sources)
            with st.expander(expander_title, expanded=True):
                # For the "1. Data Dictionary" sheet: display only required columns.
                if sheet_name == "1. Data Dictionary":
                    required_cols = [
                        "CorporateFinanceSubmissionFieldName",
                        "Corporate Finance Submission Field Description",
                        "Transformation/Business Logic"
                    ]
                    # Create a DataFrame from the rows.
                    df_all = pd.DataFrame(rows)
                    # Drop completely empty columns.
                    df_all = df_all.dropna(axis=1, how='all')
                    try:
                        # Select only the required columns plus Source.
                        df_display = df_all[["Source", 
                                             "CorporateFinanceSubmissionFieldName", 
                                             "Corporate Finance Submission Field Description", 
                                             "Transformation/Business Logic"]]
                    except KeyError:
                        st.warning("Required columns not found in some rows.")
                        df_display = pd.DataFrame(rows)
                    # Rename header for the second column to include a line break.
                    df_display = df_display.rename(
                        columns={"Corporate Finance Submission Field Description":
                                 "Corporate Finance Submission<br>Field Description"})
                    # Apply wrapping: if any cell's text exceeds 100 characters, insert <br> every 100 characters.
                    for col in df_display.columns:
                        df_display[col] = df_display[col].apply(
                            lambda x: wrap_text_html(str(x), width=100) if isinstance(x, str) and len(x) > 100 else x)
                    # Generate HTML table.
                    html_table = generate_html_table_from_df(df_display, wrap_width=100)
                    components.html(html_table, height=700, scrolling=True)
                else:
                    # For other sheets, display all non-empty columns.
                    df_display = pd.DataFrame(rows)
                    df_display = df_display.dropna(axis=1, how='all')
                    html_table = generate_html_table_from_df(df_display, wrap_width=100)
                    components.html(html_table, height=700, scrolling=True)

if __name__ == "__main__":
    main()