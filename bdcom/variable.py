import pandas as pd
import re

def extract_sas_variables(sas_code):
    """
    Extract SAS variable names from a given SAS code string, excluding comments,
    literals, date constants, format specifiers, and macro variables.
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

    # Remove date literals like '01Jan1900'D (d literal)
    cleaned_code = re.sub(r'\'[^\']*\'D', '', cleaned_code)  # Remove date literals like '01Jan1900'D
    
    # Remove format specifiers like yymmddn8. (variables with a period at the end)
    # This will remove variables like 'varname.' or 'yymmddn8.'
    cleaned_code = re.sub(r'\s+[A-Za-z_][A-Za-z0-9_]*\.\s*', '', cleaned_code)  # Remove format specifiers like yymmddn8.

    # Remove macro variables like &enddt (anything starting with &)
    cleaned_code = re.sub(r'\&\w+', '', cleaned_code)  # Remove macro variables like &enddt

    # Extract potential SAS identifiers (variable names)
    tokens = re.findall(r'\b[A-Za-z_][A-Za-z0-9_]*\b', cleaned_code)
    
    # Exclusions: SAS reserved keywords, functions, etc.
    sas_exclusions = {
        "DATA", "SET", "MERGE", "UPDATE", "BY", "IF", "THEN", "ELSE", "ELSEIF", "DO", "END", "OUTPUT",
        "DROP", "KEEP", "RENAME", "LABEL", "FORMAT", "INFORMAT", "LENGTH", "ATTRIB", "ARRAY", "RETAIN",
        "RUN", "QUIT", "LIBNAME", "FILENAME", "OPTIONS", "TITLE", "FOOTNOTE", "CARDS", "DATALINES",
        "_NULL_", "_ALL_", "_NUMERIC_", "_CHARACTER_", "INPUT", "PUT", "INFILE", "FILE", "SELECT", "FROM",
        # Add additional exclusions as needed.
    }
    
    # Return only tokens that are not excluded
    variables = {token for token in tokens if token.upper() not in sas_exclusions}
    
    return variables

def update_sql_select_clause(sql_code, variables_to_add):
    """
    Update the SQL SELECT clause with the missing variables and format the SQL.
    """
    # Find the SELECT clause and split it at "SELECT" and "FROM"
    select_start = sql_code.lower().find('select')
    from_start = sql_code.lower().find('from')
    
    if select_start == -1 or from_start == -1:
        print(f"Could not find SELECT or FROM in SQL: {sql_code}")
        return sql_code  # If SELECT or FROM is not found, return the code unchanged

    # Extract the SELECT clause from the SQL
    select_clause = sql_code[select_start + 6:from_start].strip()  # Remove 'SELECT' and get the fields
    select_clause = select_clause.replace('\n', ' ').replace('\r', ' ')  # Normalize line breaks and spaces

    # Split the existing fields in SELECT clause and convert to lowercase for case-insensitive comparison
    existing_vars = {var.strip() for var in select_clause.split(',')}
    
    # Add the new variables, avoiding duplicates (case-insensitive check)
    new_vars = [var for var in variables_to_add if var.lower() not in existing_vars]
    
    if new_vars:
        # Add missing variables to the SELECT clause
        select_clause += ", " + ", ".join(new_vars)

    # Format the SELECT clause so each variable appears on a new line
    select_clause_formatted = "SELECT\n" + ",\n".join(f"  {var}" for var in sorted(existing_vars | set(new_vars))) + "\n"

    # Ensure there's a space between the last variable and the "FROM"
    sql_code = sql_code[:select_start + 6] + select_clause_formatted + sql_code[from_start:]
    
    return sql_code

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
        
        # Add missing variables to SELECT clause
        updated_sql = update_sql_select_clause(updated_sql, combined_vars)
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
output_filename = 'updated_sql_file_with_variables_v13.xlsx'
with pd.ExcelWriter(output_filename) as writer:
    sql_df.to_excel(writer, sheet_name='Data', index=False)

print(f"Updated Excel file saved as '{output_filename}'.")