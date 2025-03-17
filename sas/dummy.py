import pandas as pd
import numpy as np
import saspy

# -----------------------------
# Step 1: Generate the dummy dataset in Python
# -----------------------------
# Create a DataFrame with 1,000,000 rows and 200 columns filled with random numbers.
n_rows = 1_000_000
n_cols = 200
# Generate column names: var1, var2, ..., var200
columns = [f'var{i}' for i in range(1, n_cols + 1)]
# Create the DataFrame using random numbers (from a normal distribution)
df = pd.DataFrame(np.random.randn(n_rows, n_cols), columns=columns)

# -----------------------------
# Step 2: Establish a SAS session using SASPy
# -----------------------------
# Ensure your SASPy configuration (e.g., cfgname 'default') is properly set up.
sas = saspy.SASsession(cfgname='default')

# Transfer the Pandas DataFrame to the SAS WORK library
# This will create a temporary SAS dataset named 'dummy_data' in the WORK library.
sas_df = sas.df2sd(df, table='dummy_data', libref='work')

# -----------------------------
# Step 3: Save the dataset as a SAS7BDAT file
# -----------------------------
# Define a libname pointing to a directory where you have write permissions.
# SAS stores datasets in SAS7BDAT format by default.
# Replace '/path/to/output/directory' with your desired output directory.
sas_code = """
libname mylib '/path/to/output/directory';
data mylib.dummy_data;
   set work.dummy_data;
run;
"""
# Submit the SAS code to save the dataset.
sas.submit(sas_code)

# (Optional) List the SAS log to verify that the dataset was written successfully.
print(sas.lst)