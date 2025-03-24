import re

def extract_sas_variables(sas_code):
    # Step 1: Remove quoted strings (handles standard and smart quotes)
    # This regex removes any text enclosed by any of these quote characters:
    #   "  '  ‘  ’  “  ”
    cleaned_code = re.sub(r'["\'\u2018\u2019\u201C\u201D].*?["\'\u2018\u2019\u201C\u201D]', '', sas_code)
    
    # Step 2: Extract tokens that look like SAS identifiers.
    # SAS variable names usually start with a letter or underscore and can contain letters, digits, or underscores.
    tokens = re.findall(r'\b[A-Za-z_][A-Za-z0-9_]*\b', cleaned_code)
    
    # Step 3: Exclude known SAS reserved words, built-in functions, macro keywords, automatic variables,
    # procedure names/options, and common logical operators.
    sas_exclusions = {
        # SAS Reserved Keywords
        "DATA", "SET", "MERGE", "UPDATE", "BY", "IF", "THEN", "ELSE", "ELSEIF", "DO", "END", "OUTPUT",
        "DROP", "KEEP", "RENAME", "LABEL", "FORMAT", "INFORMAT", "LENGTH", "ATTRIB", "ARRAY", "RETAIN",
        "RUN", "QUIT", "LIBNAME", "FILENAME", "OPTIONS", "TITLE", "FOOTNOTE", "CARDS", "DATALINES",
        "CARDS4", "DATALINES4", "_NULL_", "_ALL_", "_NUMERIC_", "_CHARACTER_", "INPUT", "PUT", "INFILE", 
        "FILE", "SELECT", "FROM", "WHERE", "GROUP", "HAVING", "ORDER", "CASE", "WHEN", "UNION", "ALL", 
        "EXAMINE", "DEFINE", "OBS", "FIRSTOBS",
        
        # SAS Built-in Functions
        "ABS", "ACOS", "ACOSH", "ALPHA", "ANYALNUM", "ANYALPHA", "ANYCNTRL", "ANYDIGIT", "ANYLOWER", 
        "ANYUPPER", "ANYSPACE", "ANYXDIGIT", "ARCCOS", "ARCCOSH", "ARSIN", "ARSINH", "ARTAN", "ARTANH", 
        "ATAN", "ATAN2", "BQUOTE", "BYTE", "CEIL", "CEILING", "COS", "COSH", "DATE", "DAY", "DECODE", 
        "DIGIT", "FLOOR", "INDEX", "INDEXC", "INDEXW", "INPUTN", "INT", "INTCK", "INTNX", "LAG", "LBOUND", 
        "LENGTH", "LOG", "LOG10", "MAX", "MEAN", "MEDIAN", "MIN", "MOD", "NVALID", "NOW", "NLOGB", "NROOT", 
        "NVAR", "PV", "RANUNI", "ROUND", "SIGN", "SIN", "SINH", "SQRT", "STRIP", "SUBSTR", "SUM", "TAN", 
        "TANH", "TRANWRD", "TRIM", "UPCASE", "LOWCASE", "COMPRESS", "COMPBL", "CAT", "CATS", "CATX", "FIND", 
        "FINDC", "FINDW", "VERIFY", "COALESECEC", "COALESCE",
        
        # SAS Macro Keywords and Functions
        "%MACRO", "%MEND", "%LET", "%IF", "%THEN", "%ELSE", "%DO", "%END", "%GOTO", "%RETURN", "%ABORT", 
        "%PUT", "%GLOBAL", "%LOCAL", "%SYMDEL", "%INCLUDE", "%WINDOW", "%DISPLAY", "%INPUT", "%EVAL", 
        "%SYSEVALF", "%SCAN", "%QSCAN", "%SUBSTR", "%QSUBSTR", "%INDEX", "%LENGTH", "%STR", "%NRSTR", 
        "%QUOTE", "%NRQUOTE", "%BQUOTE", "%NRBQUOTE", "%SUPERQ", "%SYSFUNC", "%QSYSFUNC", "%CMPRES", 
        "%QCMPRES", "%QUPCASE", "SYMPUT", "SYMPUTX", "SYMGET",
        
        # Automatic Variables (DATA Step)
        "_N_", "_ERROR_", "_FILE_", "_INFILE_", "_IORC_", "_MSG_", "_CMD_",
        
        # SAS Automatic Macro Variables (SYS variables)
        "SYSDATE", "SYSDATE9", "SYSDAY", "SYSTIME", "SYSCC", "SYSERR", "SYSFILRC", "SYSLIBRC", 
        "SYSINFO", "SYSWARNINGTEXT", "SYSERRORTEXT", "SYSLAST", "SYSDBRC", "SYSDBMSG", "SQLOBS", 
        "SQLRC", "SYSPARM", "SYSJOBID", "SYSPROCESSNAME", "SYSPROCESSID", "SYSSCP", "SYSSCPL",
        
        # SAS Procedure Names and Common Options
        "PROC", "PRINT", "SORT", "MEANS", "SUMMARY", "FREQ", "TABULATE", "REPORT", "DATASETS", "SQL", 
        "REG", "GLM", "ANOVA", "LOGISTIC", "GENMOD", "MIXED", "UNIVARIATE", "CORR", "NPAR1WAY", 
        "LIFETEST", "PHREG", "SURVEYMEANS", "SURVEYFREQ", "SURVEYLOGISTIC", "GPLOT", "GCHART", 
        "GREPLAY", "ARIMA", "AUTOREG", "EXPAND", "TIMESERIES", "ACECLUS", "CLUSTER", "FASTCLUS", 
        "VARCLUS", "PRINCOMP", "DMDB", "HPFOREST", "IML", "NLIN", "QLIM", "KRIGE2D", "GLIMMIX", "PLM", 
        "POWER", "ICPHREG", "MCMC", "STDIZE", "TRANSREG", "X12", "OUT", "NOPRINT", "NOOBS", "BY", 
        "WHERE", "PLOTS", "ALPHA", "MAXDEC", "NWAY", "ORDER",
        
        # Logical Operators and Other Reserved Words
        "NOT", "IN", "AND", "OR", "XOR", "EQ", "NE", "GT", "LT", "GE", "LE",
        
        # Additional Tokens (as encountered in your code)
        "XX"
    }
    
    # Filter tokens by removing any that match the exclusion list (using case-insensitive matching)
    variables = {token for token in tokens if token.upper() not in sas_exclusions}
    
    return variables

# Example SAS code snippet
sas_code = '''
IF UPCASE(COALESCEC(product_cd_vo, product_cd_m_s)) = ‘LOAN’
Then propstate = UPCASE(COMPBL(STRIP(property_state_cd)));
Else If COMPRESS(property_state_cd_vo) Not In(", XX")
Then propstate = UPCASE(COMPBL(STRIP(property_state_cd_vo)));
Else If COMPRESS(propstate_pm) Not In(", ‘LOAN’")
Then propstate = UPCASE(COMPBL(STRIP(propstate_pm)));
Else propstate = "";
AND;
'''

extracted_vars = extract_sas_variables(sas_code)
print("Extracted Variables:", extracted_vars)