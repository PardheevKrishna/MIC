import pandas as pd
import re

def extract_sas_variables(sas_code):
    """
    Extract SAS variable names from a given SAS code string.
    Non-string inputs return an empty set.
    """
    if not isinstance(sas_code, str):
        return set()
    
    # Remove various quoted strings (standard and smart quotes)
    cleaned_code = re.sub(r'["\'].+?["\']', '', sas_code)
    cleaned_code = re.sub(r'[\u2018].+?[\u2019]', '', cleaned_code)
    cleaned_code = re.sub(r'[\u201C].+?[\u201D]', '', cleaned_code)
    
    # Extract potential SAS identifiers
    tokens = re.findall(r'\b[A-Za-z_][A-Za-z0-9_]*\b', cleaned_code)
    
    # Exclude common SAS reserved words, functions, etc.
    sas_exclusions = {
        "DATA", "SET", "MERGE", "UPDATE", "BY", "IF", "THEN", "ELSE", "ELSEIF", "DO", "END", "OUTPUT",
        "DROP", "KEEP", "RENAME", "LABEL", "FORMAT", "INFORMAT", "LENGTH", "ATTRIB", "ARRAY", "RETAIN",
        "RUN", "QUIT", "LIBNAME", "FILENAME", "OPTIONS", "TITLE", "FOOTNOTE", "CARDS", "DATALINES",
        "_NULL_", "_ALL_", "_NUMERIC_", "_CHARACTER_", "INPUT", "PUT", "INFILE", "FILE", "SELECT", "FROM",
        # Add more exclusions as needed...
    }
    
    variables = {token for token in tokens if token.upper() not in sas_exclusions}
    return variables

# Load the Excel files (ensure the file paths are correct)
derivation_df = pd.read_excel('data derivation.xlsx', sheet_name='2.  First Mortgage File')
sql_df = pd.read_excel('myexcel.xlsx', sheet_name='Data')

# Group rows by business name (the variable name) so we can concatenate SAS code and collect source fields
variable_to_source_map = {}

# Group by the business name column; drop rows where the business name is missing
grouped = derivation_df.groupby('Variable Name\n(Business Name)', dropna=True)

for var_name, group in grouped:
    # Concatenate all SAS code rows for this business name into one block
    sas_code_block = " ".join(group['Logic to Populate FR Y-14M Field'].astype(str).tolist())
    sas_vars = extract_sas_variables(sas_code_block)
    
    # Collect all manually provided source fields from each row; skip non-string values
    source_fields_set = set()
    for val in group['CLRTY/Source Fields Used']:
        if isinstance(val, str):
            # In case there are multiple lines in one cell, split them; otherwise, add the value directly.
            for field in val.splitlines():
                field = field.strip()
                if field:
                    source_fields_set.add(field)
    
    # Combine the extracted SAS variables with the manually provided source fields
    combined_vars = source_fields_set | sas_vars
    variable_to_source_map[var_name] = combined_vars

# Now update the SQL logic in the SQL file based on the variable-to-source mapping.
updated_sql_logic = []
for i, row in sql_df.iterrows():
    sql_code = row['value_sql_logic']
    field_name = row['field_name']
    
    # Replace escape characters for proper formatting
    sql_code = sql_code.replace(r'\r', ' ')
    sql_code = sql_code.replace(r'\t', ' ')
    sql_code = sql_code.replace(r'\n', '\n')
    
    # If a matching business name exists, append its source variables to the SELECT clause.
    # Here we assume that the business name (field_name) in myexcel.xlsx is in uppercase with no spaces.
    if field_name in variable_to_source_map:
        select_vars = "\n".join(f"  {var}," for var in variable_to_source_map[field_name])
        # Insert the source variables right after the SELECT keyword.
        sql_code = sql_code.replace("SELECT", f"SELECT\n{select_vars}", 1)
    
    updated_sql_logic.append(sql_code)

# Update the DataFrame and save to a new Excel file.
sql_df['value_sql_logic'] = updated_sql_logic

with pd.ExcelWriter('updated_sql_file.xlsx') as writer:
    sql_df.to_excel(writer, sheet_name='Data', index=False)

print("SQL logic updated and saved to 'updated_sql_file.xlsx'.")