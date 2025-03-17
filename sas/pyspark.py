# test_saspy.py

import saspy

# Create a SAS session using the 'default' config in sascfg_personal.py
sas = saspy.SASsession(cfgname='default')

# Print the SAS version info to verify the connection
print("SAS Session Initialized.")
print("SAS Version Info:")
print(sas.sasproductlevels())

# Run a test procedure (PROC MEANS) on sashelp.class
print("\nRunning PROC MEANS on sashelp.class...")
result = sas.submit("""
proc means data=sashelp.class;
run;
""")

# Print the LOG and LST output
print("\n--- SAS LOG ---")
print(result['LOG'])

print("\n--- SAS LISTING ---")
print(result['LST'])

# End the SAS session
sas._endsas()