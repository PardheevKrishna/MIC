import logging
import numpy as np
import pandas as pd
import saspy
from time import time

###############################################################################
# 1. SASPy CONFIGURATION (IOM on Windows)
###############################################################################
# We'll define a configuration dictionary inline and tell SASPy to use it.
# Update the paths to match your local environment!

my_configs = {
    'SAS_config_names': ['winlocal'],

    'winlocal': {
        # Path to your local Java runtime
        'java': r'C:\Program Files\Java\jre1.8.0_251\bin\java.exe',  # <-- Update!

        # For a local SAS install on Windows, usually 'localhost' and port 8591
        'iomhost': 'localhost',
        'iomport': 8591,

        # Classpath entries pointing to the JAR files in your SASPy 'java' directory
        'classpath': [
            r'C:\path\to\saspy\java\sas.core.jar',     # <-- Update!
            r'C:\path\to\saspy\java\saspyiom.jar',     # <-- Update!
            r'C:\path\to\saspy\java\log4j.jar'         # <-- Update!
        ],

        # Common encoding on Windows
        'encoding': 'windows-1252',

        # Increase if needed for large data
        'timeout': 9999
    }
}

###############################################################################
# 2. LOGGING SETUP
###############################################################################
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

###############################################################################
# 3. START THE SAS SESSION
###############################################################################
# Pass the config dictionary and the name of the config we want to use
logging.debug("Starting SAS session via SASPy (IOM on Windows).")
sas = saspy.SASsession(cfgdict=my_configs, cfgname='winlocal')
logging.debug("SAS session started successfully (if no errors above).")

###############################################################################
# 4. GENERATE DUMMY DATA (1M rows, 200 columns)
###############################################################################
# Let's split them into 100 numerical and 100 categorical for variety.
num_rows = 1_000_000
num_numerical = 100
num_categorical = 100
chunk_size = 100_000
categories = ['A', 'B', 'C', 'D', 'E']

logging.debug("Generating dummy data: 1 million rows, 200 columns in chunks.")

df_chunks = []
num_chunks = num_rows // chunk_size

start_time = time()
for chunk_index in range(num_chunks):
    logging.debug(f"Generating chunk {chunk_index+1}/{num_chunks}")
    # Numerical data
    numerical_data = np.random.rand(chunk_size, num_numerical)
    # Categorical data
    categorical_data = np.random.choice(categories, size=(chunk_size, num_categorical))

    # Build chunk DataFrames
    df_num = pd.DataFrame(numerical_data, columns=[f'num_{i+1}' for i in range(num_numerical)])
    df_cat = pd.DataFrame(categorical_data, columns=[f'cat_{i+1}' for i in range(num_categorical)])
    df_chunk = pd.concat([df_num, df_cat], axis=1)
    df_chunks.append(df_chunk)

# Handle remainder if rows not multiple of chunk_size
remainder = num_rows % chunk_size
if remainder > 0:
    logging.debug(f"Generating remainder chunk with {remainder} rows.")
    numerical_data = np.random.rand(remainder, num_numerical)
    categorical_data = np.random.choice(categories, size=(remainder, num_categorical))

    df_num = pd.DataFrame(numerical_data, columns=[f'num_{i+1}' for i in range(num_numerical)])
    df_cat = pd.DataFrame(categorical_data, columns=[f'cat_{i+1}' for i in range(num_categorical)])
    df_chunk = pd.concat([df_num, df_cat], axis=1)
    df_chunks.append(df_chunk)

# Concatenate all chunks
df = pd.concat(df_chunks, ignore_index=True)
logging.debug(f"Dummy data generated. Shape: {df.shape}")

end_time = time()
logging.debug(f"Data generation took {end_time - start_time:.2f} seconds.")

###############################################################################
# 5. TRANSFER THE DATAFRAME TO SAS (WORK LIBRARY)
###############################################################################
logging.debug("Transferring DataFrame to SAS WORK library as 'dummy_data'.")
sas.df2sd(df, table='dummy_data', libref='work')
logging.debug("Transfer complete.")

###############################################################################
# 6. SAVE DATA PERMANENTLY AS A .sas7bdat FILE
###############################################################################
# We'll define a library pointing to a directory on your machine. Update the path!
libref = 'mylib'
output_path = r'C:\path\to\output'  # <-- Change to a valid directory for SAS
sas_code = f"""
libname {libref} '{output_path}';
data {libref}.dummy_data;
    set work.dummy_data;
run;
"""
logging.debug("Submitting SAS code to define library and copy dataset.")
sas.submit(sas_code)
logging.debug("SAS dataset creation complete.")

print("Data exported as a native SAS dataset (.sas7bdat) successfully.")