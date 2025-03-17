# sascfg_personal.py
# ----------------------------------------------------
# Minimal saspy configuration for local Windows SAS
# ----------------------------------------------------

SAS_config_names = ['default']

default = {
    # Path to the SAS executable (64-bit). Adjust as needed.
    # Double up the backslashes or use raw strings in Python.
    'saspath': 'C:\\Program Files\\SASHome\\SASFoundation\\9.4\\sas.exe',
    
    # Standard options for a no-interactive SAS session
    'options': ['-nosplash', '-sysin', ' '],
    
    # Encoding for Windows SAS sessions (often 'windows-1252')
    'encoding': 'windows-1252',
    
    # Run SAS locally on Windows
    'mode': 'local',
    'host': 'winlocal'
}