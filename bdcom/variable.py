import pandas as pd
import re

def extract_sas_variables(sas_code):
    """
    Extract SAS variable names from a given SAS code string, excluding comments.
    """
    if not isinstance(sas_code, str):
        return set()

    # Remove block comments (/* ... */) and line comments (// or * ... ;) from SAS code
    sas_code_no_comments = re.sub(r'/\*.*?\*/', '', sas_code, flags=re.DOTALL)  # Remove block comments
    sas_code_no_comments = re.sub(r'(?<=\n)\*.*?;', '', sas_code_no_comments)  # Remove line comments

    # Remove quoted strings (standard and smart quotes)
    cleaned_code = re.sub(r'["\'].+?["\']', '', sas_code_no_comments)
    cleaned_code = re.sub(r'[\u2018].+?[\u2019]', '', cleaned_code)
    cleaned_code = re.sub(r'[\u201C].+?[\u201D]', '', cleaned_code)

    # Extract potential SAS identifiers
    tokens = re.findall(r'\b[A-Za-z_][A-Za-z0-9_]*\b', cleaned_code)
    
    # Exclusions: SAS reserved keywords, functions, etc.
    sas_exclusions = {
        "DATA", "SET", "MERGE", "UPDATE", "BY", "IF", "THEN", "ELSE", "ELSEIF", "DO", "END", "OUTPUT",
        "DROP", "KEEP", "RENAME", "LABEL", "FORMAT", "INFORMAT", "LENGTH", "ATTRIB", "ARRAY", "RETAIN",
        "RUN", "QUIT", "LIBNAME", "FILENAME", "OPTIONS", "TITLE", "FOOTNOTE", "CARDS", "DATALINES",
        "_NULL_", "_ALL_", "_NUMERIC_", "_CHARACTER_", "INPUT", "PUT", "INFILE", "FILE", "SELECT", "FROM",
        # Add additional exclusions as needed.
    }
    
    variables = {token for token in tokens if token.upper() not in sas_exclusions}
    
    return variables

def update_sql_with_variables(sql_code, extracted_vars):
    """
    Update the SQL SELECT clause with the extracted variables, ensuring no duplicates.
    """
    # Improved regex to handle multi-line SQL with flexible spacing
    select_clause_match = re.search(r"SELECT\s+(.+?)\s+FROM", sql_code, re.DOTALL)
    
    # Print the SQL for debugging
    if select_clause_match is None:
        print("No match for SELECT clause in SQL code:\n", sql_code)  # Debug print
        return sql_code  # If SELECT is not found, return the code unchanged
    
    # Extract existing variables in SELECT clause and clean them up
    existing_vars_str = select_clause_match.group(1).strip()
    existing_vars = {var.strip() for var in existing_vars_str.split(',')}
    
    # Add the new variables, avoiding duplicates
    all_vars = existing_vars | extracted_vars  # Combine existing and new variables
    
    # Create the new SELECT clause
    new_select_clause = "SELECT\n" + existing_vars_str + ",\n" + ",\n".join(f"  {var}" for var in sorted(all_vars)) + "\n"

    # Replace the old SELECT clause with the new one
    updated_sql = re.sub(r"SELECT\s+(.+?)\s+FROM", new_select_clause + "FROM", sql_code, flags=re.DOTALL)

    return updated_sql

# Load Excel files (ensure the file paths are correct)
derivation_df = pd.read_excel('data derivation.xlsx', sheet_name='2.  First Mortgage File')
sql_df = pd.read_excel('myexcel.xlsx', sheet_name='Data')

# Create dictionaries to store for each business name:
# - The concatenated SAS code
# - The union of manually provided source fields and SAS variables
normalized_map = {}

# Group the derivation file by the business name column.
# We drop rows where the business name is missing.
grouped = derivation_df.groupby('Variable Name\n(Business Name)', dropna=True)

for business_name, group in grouped:
    # Concatenate all SAS code rows for this business name.
    sas_code_block = " ".join(group['Logic to Populate FR Y-14M Field'].astype(str).tolist())
    sas_vars = extract_sas_variables(sas_code_block)
    
    # Collect all manually provided source fields from the group.
    source_fields_set = set()
    for val in group['CLRTY/Source Fields Used']:
        if isinstance(val, str):
            # In case a cell has multiple lines, split them.
            for field in val.splitlines():
                field = field.strip()
                if field:
                    source_fields_set.add(field)
    
    # Union the manually provided fields with the extracted SAS variables.
    combined_vars = source_fields_set | sas_vars
    
    # Create a normalized key to match with SQL file field names (uppercase, no spaces)
    norm_key = business_name.replace(" ", "").upper()
    normalized_map[norm_key] = {
        'combined_vars': combined_vars,
        'sas_code': sas_code_block
    }

# Prepare lists for the new columns
all_var_names_list = []
old_sql_code_list = []
new_sql_code_list = []
sas_code_list = []

# Process each row in the SQL file to update SQL code and add extra columns.
for idx, row in sql_df.iterrows():
    original_sql = row['value_sql_logic']
    # Format the original SQL code by replacing escape sequences.
    formatted_old_sql = original_sql.replace(r'\r', ' ').replace(r'\t', ' ').replace(r'\n', '\n')
    
    field_name = row['field_name']  # Assumed to be in uppercase with no spaces.
    # Initialize new SQL code as the original formatted code.
    updated_sql = formatted_old_sql
    
    if field_name in normalized_map:
        # Get combined variables and sas code from the mapping.
        combined_vars = normalized_map[field_name]['combined_vars']
        sas_code_for_field = normalized_map[field_name]['sas_code']
        
        # Update the SQL with the extracted variables (if they are not already in SELECT)
        updated_sql = update_sql_with_variables(updated_sql, combined_vars)
    else:
        combined_vars = []
        sas_code_for_field = ""
    
    # Append the details to the corresponding lists.
    all_var_names_list.append(", ".join(sorted(combined_vars)))
    old_sql_code_list.append(formatted_old_sql)
    new_sql_code_list.append(updated_sql)
    sas_code_list.append(sas_code_for_field)

# Add the new columns to the SQL DataFrame.
sql_df['all_variable_names'] = all_var_names_list
sql_df['old_sql_code'] = old_sql_code_list
sql_df['new_sql_code'] = new_sql_code_list
sql_df['sas_code'] = sas_code_list

# Save the updated DataFrame to a new Excel file.
output_filename = 'updated_sql_file_with_variables_v9.xlsx'
with pd.ExcelWriter(output_filename) as writer:
    sql_df.to_excel(writer, sheet_name='Data', index=False)

print(f"Updated Excel file saved as '{output_filename}'.")