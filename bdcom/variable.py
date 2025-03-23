import re

def extract_sas_variables(sas_code):
    # Step 1: Remove quoted strings (both single and double quotes)
    cleaned_code = re.sub(r'(["\']).*?\1', '', sas_code)
    
    # Step 2: Extract tokens that look like SAS identifiers.
    # SAS variable names usually start with a letter or underscore and can contain letters, digits, or underscores.
    tokens = re.findall(r'\b[A-Za-z_][A-Za-z0-9_]*\b', cleaned_code)
    
    # Step 3: Exclude known SAS keywords and function names.
    # You can extend this set based on your SAS environment.
    sas_exclusions = {
        'IF', 'THEN', 'ELSE', 'ELSEIF', 'NOT', 'IN', 
        'COMPRESS', 'COMPBL', 'STRIP', 'UPCASE', 'DO', 'END'
    }
    
    # Filter tokens: compare in uppercase for case-insensitive matching
    variables = {token for token in tokens if token.upper() not in sas_exclusions}
    
    return variables

# Example SAS code snippet
sas_code = '''
IF COMPRESS(property_state_cd) Not In(", XX')
Then propstate = UPCASE(COMPBL(STRIP(property_state_cd)));
Else If COMPRESS(property_state_cd_vo) Not In(", XX)
Then propstate = UPCASE(COMPBL(STRIP(property_state_cd_vo)));
Else If COMPRESS(propstate_pm) Not In(", 'XX')
Then propstate = UPCASE(COMPBL(STRIP(propstate_pm)));
Else propstate = ";
'''

extracted_vars = extract_sas_variables(sas_code)
print("Extracted Variables:", extracted_vars)