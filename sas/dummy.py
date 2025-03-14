import pandas as pd
import numpy as np
import logging
from time import time

# Configure logging to display debug messages with timestamps.
logging.basicConfig(level=logging.DEBUG,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Parameters for dummy data generation
num_rows = 1_000_000
num_numerical = 100       # Number of numerical columns
num_categorical = 100     # Number of categorical columns
chunk_size = 100_000      # Number of rows per chunk
num_chunks = num_rows // chunk_size
categories = ['A', 'B', 'C', 'D', 'E']  # List of possible categories

logging.debug(f"Starting dummy data generation: {num_rows} rows, {num_numerical} numerical and {num_categorical} categorical columns in {num_chunks} chunks.")

# List to store DataFrame chunks
df_chunks = []
start_time = time()

# Generate data in chunks
for chunk in range(num_chunks):
    logging.debug(f"Generating chunk {chunk + 1}/{num_chunks}")
    
    # Create numerical data: random floats between 0 and 1
    numerical_data = np.random.rand(chunk_size, num_numerical)
    
    # Create categorical data: random choices from the specified list
    categorical_data = np.random.choice(categories, size=(chunk_size, num_categorical))
    
    # Create DataFrames for each type of data with appropriate column names
    df_numerical = pd.DataFrame(numerical_data, columns=[f'num_{i+1}' for i in range(num_numerical)])
    df_categorical = pd.DataFrame(categorical_data, columns=[f'cat_{i+1}' for i in range(num_categorical)])
    
    # Concatenate numerical and categorical data side-by-side
    df_chunk = pd.concat([df_numerical, df_categorical], axis=1)
    df_chunks.append(df_chunk)
    
    logging.debug(f"Chunk {chunk + 1} generated.")

# If there are any remaining rows (if num_rows is not an exact multiple of chunk_size)
remainder = num_rows % chunk_size
if remainder:
    logging.debug(f"Generating remainder chunk with {remainder} rows")
    numerical_data = np.random.rand(remainder, num_numerical)
    categorical_data = np.random.choice(categories, size=(remainder, num_categorical))
    df_numerical = pd.DataFrame(numerical_data, columns=[f'num_{i+1}' for i in range(num_numerical)])
    df_categorical = pd.DataFrame(categorical_data, columns=[f'cat_{i+1}' for i in range(num_categorical)])
    df_chunk = pd.concat([df_numerical, df_categorical], axis=1)
    df_chunks.append(df_chunk)
    logging.debug("Remainder chunk generated.")

# Concatenate all chunks into one DataFrame
df = pd.concat(df_chunks, ignore_index=True)
end_time = time()

logging.debug(f"Data generation complete. Total time: {end_time - start_time:.2f} seconds")
print("Dummy data generation complete. DataFrame shape:", df.shape)