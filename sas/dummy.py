import pandas as pd
import numpy as np
import logging
from time import time
import pyreadstat

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
              f"({num_numerical} numerical, {num_categorical} categorical) "
              f"in chunks of {chunk_size} rows.")

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
    
    # Create DataFrames for numerical and categorical data
    df_numerical = pd.DataFrame(numerical_data, columns=[f'num_{i+1}' for i in range(num_numerical)])
    df_categorical = pd.DataFrame(categorical_data, columns=[f'cat_{i+1}' for i in range(num_categorical)])
    
    # Concatenate the two DataFrames side by side
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
logging.debug(f"Dummy data generation complete. Total time: {end_time - start_time:.2f} seconds.")

###############################################################################
# EXPORT EACH CHUNK TO A SAS XPORT FILE AND UPDATE PROGRESS
###############################################################################
total_rows_written = 0
for i, df_chunk in enumerate(df_chunks):
    chunk_filename = f"dummy_data_chunk_{i+1}.xpt"
    pyreadstat.write_xport(df_chunk, chunk_filename)
    
    total_rows_written += df_chunk.shape[0]
    logging.debug(f"Chunk {i+1} written to {chunk_filename}. Total rows written: {total_rows_written} / {num_rows}")
    print(f"Progress: {total_rows_written} / {num_rows} rows written.")

print("All chunks have been written.")
print("Exported files: dummy_data_chunk_*.xpt")