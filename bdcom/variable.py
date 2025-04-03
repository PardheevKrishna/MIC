import pandas as pd
import re

# Function to extract SAS variables from SAS code
def extract_sas_variables(sas_code):
    # Check if sas_code is a string. If it's not, return an empty set
    if not isinstance(sas_code, str):
        return set()

    # Remove quoted strings
    cleaned_code = re.sub(r'["\'].+?["\']', '', sas_code)
    cleaned_code = re.sub(r'[\u2018].+?[\u2019]', '', cleaned_code)
    cleaned_code = re.sub(r'[\u201C].+?[\u201D]', '', cleaned_code)
    
    # Extract SAS variable names
    tokens = re.findall(r'\b[A-Za-z_][A-Za-z0-9_]*\b', cleaned_code)
    
    # SAS reserved keywords and functions
    sas_exclusions = {
        "DATA", "SET", "MERGE", "UPDATE", "BY", "IF", "THEN", "ELSE", "ELSEIF", "DO", "END", "OUTPUT",
        "DROP", "KEEP", "RENAME", "LABEL", "FORMAT", "INFORMAT", "LENGTH", "ATTRIB", "ARRAY", "RETAIN",
        "RUN", "QUIT", "LIBNAME", "FILENAME", "OPTIONS", "TITLE", "FOOTNOTE", "CARDS", "DATALINES",
        "_NULL_", "_ALL_", "_NUMERIC_", "_CHARACTER_", "INPUT", "PUT", "INFILE", "FILE", "SELECT", "FROM",
        # Add more exclusions as needed...
    }
    
    # Filter out exclusions
    variables = {token for token in tokens if token.upper() not in sas_exclusions}
    
    return variables

# Load both Excel files
derivation_df = pd.read_excel('data derivation.xlsx', sheet_name='2.  First Mortgage File')
sql_df = pd.read_excel('myexcel.xlsx', sheet_name='Data')

# Process the derivation sheet to extract variables and source fields
variable_to_source_map = {}

# Handle grouped SAS code (grouping by business name)
current_variable_name = None
sas_code_group = ""

# Process the rows from the derivation file
for i, row in derivation_df.iterrows():
    variable_name = row['Variable Name\n(Business Name)']
    sas_code = row['Logic to Populate FR Y-14M Field']
    source_fields = row['CLRTY/Source Fields Used'].splitlines()
    
    # If the variable name changes, process the previous block
    if variable_name != current_variable_name:
        if current_variable_name is not None:
            # Process the previous variable's group
            sas_vars = extract_sas_variables(sas_code_group)
            combined_vars = set(source_fields) | sas_vars  # Union of source fields and extracted variables
            variable_to_source_map[current_variable_name] = combined_vars
        # Reset for new variable name
        sas_code_group = sas_code
        current_variable_name = variable_name
    else:
        # If the variable name is the same, append the SAS code
        sas_code_group += " " + str(sas_code)

# Process the last group (after the loop ends)
if current_variable_name is not None:
    sas_vars = extract_sas_variables(sas_code_group)
    combined_vars = set(source_fields) | sas_vars  # Union of source fields and extracted variables
    variable_to_source_map[current_variable_name] = combined_vars

# Process the SQL logic in 'myexcel.xlsx' and update the 'value_sql_logic' with the source variables
updated_sql_logic = []

for i, row in sql_df.iterrows():
    sql_code = row['value_sql_logic']
    field_name = row['field_name']
    
    # Replace \r, \t, \n with appropriate SQL formatting
    sql_code = sql_code.replace(r'\r', ' ')  # Adjusting return characters to space
    sql_code = sql_code.replace(r'\t', ' ')  # Adjusting tab characters
    sql_code = sql_code.replace(r'\n', '\n')  # Keeping newline for better readability

    # Append the relevant source variables under the SELECT clause for the matching field_name
    if field_name in variable_to_source_map:
        new_select_clause = "\n".join(f"  {var}," for var in variable_to_source_map[field_name])
        sql_code = sql_code.replace("SELECT", f"SELECT\n{new_select_clause}", 1)
    
    updated_sql_logic.append(sql_code)

# Update the SQL dataframe with the new SQL logic
sql_df['value_sql_logic'] = updated_sql_logic

# Save the updated DataFrame into a new Excel file
with pd.ExcelWriter('updated_sql_file.xlsx') as writer:
    sql_df.to_excel(writer, sheet_name='Data', index=False)

print("SQL logic updated and saved to 'updated_sql_file.xlsx'.")