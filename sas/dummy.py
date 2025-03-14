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

logging.debug(f"Starting dummy data generation: {num_rows} rows with "
              f"{num_numerical + num_categorical} columns "
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
    
    # Concatenate numerical and categorical DataFrames side by side
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

end_time = time()
logging.debug(f"Dummy data generation complete. Total generation time: {end_time - start_time:.2f} seconds.")

###############################################################################
# COMBINE CHUNKS INTO A SINGLE DATAFRAME WITH PROGRESS UPDATES
###############################################################################
combined_chunks = []
total_rows_added = 0

logging.debug("Starting to combine chunks into a single DataFrame.")
for i, chunk in enumerate(df_chunks):
    combined_chunks.append(chunk)
    total_rows_added += chunk.shape[0]
    logging.debug(f"Added chunk {i+1}: {total_rows_added} / {num_rows} rows combined so far.")
    print(f"Progress: {total_rows_added} / {num_rows} rows combined.")

logging.debug("Concatenating all chunks into the final DataFrame.")
df_final = pd.concat(combined_chunks, ignore_index=True)
logging.debug(f"Final DataFrame shape: {df_final.shape}")

###############################################################################
# WRITE THE FINAL DATAFRAME TO A SINGLE SAS XPORT FILE WITH A SPINNER
###############################################################################
output_filename = 'dummy_data.xpt'
logging.debug(f"Starting to write final DataFrame to SAS XPORT file: {output_filename}")

progress_done = False

def spinner():
    spinner_chars = ['-', '\\', '|', '/']
    i = 0
    while not progress_done:
        print(f"Writing to XPT file... {spinner_chars[i % len(spinner_chars)]}\r", end="", flush=True)
        i += 1
        sleep(0.5)
    print("Writing to XPT file... done!         ")

spinner_thread = threading.Thread(target=spinner)
spinner_thread.start()

# Write the complete DataFrame to a single XPT file (this is a blocking call)
pyreadstat.write_xport(df_final, output_filename)

progress_done = True
spinner_thread.join()

logging.debug("Final SAS XPORT file written successfully.")
print("Final SAS XPORT file created:", output_filename)