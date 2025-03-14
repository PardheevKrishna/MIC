import pandas as pd
import numpy as np
import logging
from time import time, sleep
import pyreadstat
import threading

# Configure logging to display debug messages with timestamps.
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

###############################################################################
# PARAMETERS FOR DUMMY DATA GENERATION
###############################################################################
num_rows = 1_000_000
num_numerical = 100       # Number of numerical columns
num_categorical = 100     # Number of categorical columns
chunk_size = 100_000      # Number of rows per chunk
categories = ['A', 'B', 'C', 'D', 'E']

logging.debug(f"Starting dummy data generation: {num_rows} rows with {num_numerical + num_categorical} columns "
              f"({num_numerical} numerical, {num_categorical} categorical) in chunks of {chunk_size} rows.")

###############################################################################
# GENERATE DUMMY DATA IN CHUNKS
###############################################################################
df_chunks = []
num_chunks = num_rows // chunk_size

start_time = time()
for chunk_index in range(num_chunks):
    logging.debug(f"Generating chunk {chunk_index + 1}/{num_chunks}")
    
    # Generate numerical data: random floats between 0 and 1
    numerical_data = np.random.rand(chunk_size, num_numerical)
    # Generate categorical data: random selections from the provided categories
    categorical_data = np.random.choice(categories, size=(chunk_size, num_categorical))
    
    # Create DataFrames for each set of data
    df_numerical = pd.DataFrame(numerical_data, columns=[f'num_{i+1}' for i in range(num_numerical)])
    df_categorical = pd.DataFrame(categorical_data, columns=[f'cat_{i+1}' for i in range(num_categorical)])
    
    # Concatenate numerical and categorical DataFrames side-by-side
    df_chunk = pd.concat([df_numerical, df_categorical], axis=1)
    df_chunks.append(df_chunk)

# Handle any remaining rows if num_rows is not an exact multiple of chunk_size
remainder = num_rows % chunk_size
if remainder > 0:
    logging.debug(f"Generating remainder chunk with {remainder} rows")
    numerical_data = np.random.rand(remainder, num_numerical)
    categorical_data = np.random.choice(categories, size=(remainder, num_categorical))
    
    df_numerical = pd.DataFrame(numerical_data, columns=[f'num_{i+1}' for i in range(num_numerical)])
    df_categorical = pd.DataFrame(categorical_data, columns=[f'cat_{i+1}' for i in range(num_categorical)])
    df_chunk = pd.concat([df_numerical, df_categorical], axis=1)
    df_chunks.append(df_chunk)

# Concatenate all chunks into one DataFrame
df = pd.concat(df_chunks, ignore_index=True)
logging.debug(f"Dummy data generation complete. DataFrame shape: {df.shape}")

end_time = time()
logging.debug(f"Data generation took {end_time - start_time:.2f} seconds.")

###############################################################################
# WRITE DATAFRAME TO A SAS XPORT FILE WITH A PROGRESS SPINNER
###############################################################################
output_filename = 'dummy_data.xpt'
logging.debug(f"Starting to write DataFrame to SAS XPORT file: {output_filename}")

# Create a flag and a spinner function to show progress during the write operation.
progress_done = False

def progress_spinner():
    spinner_chars = ['-', '\\', '|', '/']
    i = 0
    while not progress_done:
        print(f"Writing to XPT file... {spinner_chars[i % len(spinner_chars)]}\r", end="", flush=True)
        i += 1
        sleep(0.5)
    print("Writing to XPT file... done!         ")

# Start the spinner thread.
spinner_thread = threading.Thread(target=progress_spinner)
spinner_thread.start()

# Write the DataFrame to a SAS XPORT file.
pyreadstat.write_xport(df, output_filename)

# Signal the spinner to stop and wait for the thread to finish.
progress_done = True
spinner_thread.join()

logging.debug("SAS XPORT file written successfully.")
print("Dummy data generation complete. SAS XPORT file created:", output_filename)