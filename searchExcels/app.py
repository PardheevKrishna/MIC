import streamlit as st
import pandas as pd
import textwrap
from rapidfuzz import fuzz, process
import io

# Set the layout to wide.
st.set_page_config(layout="wide")

# Helper function that inserts newline characters every 'width' characters.
def wrap_text(text, width=100):
    return "\n".join(text[i:i+width] for i in range(0, len(text), width))

# Function to tokenize a string into words
def tokenize(text):
    return text.lower().split()

# Optimized Search using multiple fuzz matching techniques
def optimized_search(cell_text, search_terms):
    # Tokenize both cell and search term
    cell_tokens = tokenize(cell_text)
    
    # Check if any token from the search terms matches with any word in the cell text
    for term in search_terms:
        # Use fuzzy partial match (better for substring matches)
        partial_match = fuzz.partial_ratio(term.lower(), cell_text.lower())
        token_sort_match = fuzz.token_sort_ratio(term.lower(), ' '.join(cell_tokens))
        
        # Consider a match if either partial or token sort match is above the threshold (e.g., 80)
        if partial_match >= 80 or token_sort_match >= 80:
            return True
    return False

# Process a single Excel file.
def process_excel_file(uploaded_file, source_name, search_term):
    results = {}
    search_terms = search_term.split(",")  # Support multiple terms separated by commas
    
    try:
        # Read all sheets from the Excel file.
        excel_data = pd.read_excel(uploaded_file, sheet_name=None)
    except Exception as e:
        st.error(f"Error reading {source_name}: {e}")
        return results

    # Loop over each sheet.
    for sheet_name, df in excel_data.items():
        # Remove empty columns before displaying.
        df = df.dropna(axis=1, how='all')
        
        sheet_results = []
        # Iterate over each row.
        for idx, row in df.iterrows():
            row_matched = False
            # For each cell, check for partial or token-sort matches.
            for col in df.columns:
                cell_value = row[col]
                if pd.isna(cell_value):
                    continue
                cell_text = str(cell_value)
                if optimized_search(cell_text, search_terms):
                    row_matched = True
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
    search_term = st.text_input("Enter search term (separate with commas for multiple terms)")
    
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
                # Display each matching row separately with its respective columns
                for row in rows:
                    row_display = {col: row[col] for col in row if col != "Row Index"}
                    st.write(f"**Row {row['Row Index']}**")
                    st.dataframe(pd.DataFrame([row_display]), use_container_width=True)

if __name__ == "__main__":
    main()