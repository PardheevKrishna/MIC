import pandas as pd
import numpy as np
import logging
from time import time
import saspy

# Configure logging to display debug messages with timestamps.
logging.basicConfig(level=logging.DEBUG,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# -------------------------
# Dummy Data Generation
# -------------------------
num_rows = 1_000_000
num_numerical = 100       # Number of numerical columns
num_categorical = 100     # Number of categorical columns
categories = ['A', 'B', 'C', 'D', 'E']

logging.debug("Starting dummy data generation.")
# For demonstration, we generate the data in one go.
numerical_data = np.random.rand(num_rows, num_numerical)
categorical_data = np.random.choice(categories, size=(num_rows, num_categorical))

df_numerical = pd.DataFrame(numerical_data, columns=[f'num_{i+1}' for i in range(num_numerical)])
df_categorical = pd.DataFrame(categorical_data, columns=[f'cat_{i+1}' for i in range(num_categorical)])
df = pd.concat([df_numerical, df_categorical], axis=1)
logging.debug(f"Data generation complete. DataFrame shape: {df.shape}")

# -------------------------
# Export to SAS (sas7bdat)
# -------------------------
# Start a SAS session via SASPy. Ensure SASPy is configured with your SAS installation.
logging.debug("Starting SAS session via SASPy.")
sas = saspy.SASsession()

# Transfer the pandas DataFrame to SAS as a temporary dataset in the WORK library.
logging.debug("Converting pandas DataFrame to a SAS dataset (temporary in WORK).")
sas.df2sd(df, table='dummy_data', libref='work')

# Define a permanent SAS library.
# Update '/path/to/output' with an actual directory on your SAS server where you have write access.
libref = 'mylib'
output_path = '/path/to/output'  # <-- Change this to your desired output path.
sas.submit(f"libname {libref} '{output_path}';")

# Copy the temporary dataset (WORK.dummy_data) to the permanent library (mylib),
# which will create a sas7bdat file in the specified output directory.
logging.debug("Copying dataset from WORK to permanent SAS library to create a .sas7bdat file.")
sas.submit(f"""
data {libref}.dummy_data;
    set work.dummy_data;
run;
""")

logging.debug("SAS dataset creation complete. The .sas7bdat file should now be in the specified output directory.")
print("Data exported as a SAS dataset (.sas7bdat file) successfully.")