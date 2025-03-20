import pandas as pd
import numpy as np
import dask.dataframe as dd
from dask.diagnostics import ProgressBar
from numba import njit
import time

print("DEBUG: Starting complex processing...")
start_time = time.time()

# Define a heavy computation function with Numba acceleration
@njit
def compute_array(values1, values2):
    n = values1.shape[0]
    result = np.empty(n)
    for i in range(n):
        s = 0.0
        # Simulate heavy computation with an inner loop
        for j in range(100):
            s += np.sin(values1[i]) * np.cos(values2[i]) / (j + 1)
        result[i] = s
    return result

# Function to process each Dask partition using the numba-accelerated function
def process_partition(df):
    arr = compute_array(df['value1'].values, df['value2'].values)
    df['computed'] = arr
    # Debugging: Print number of processed rows in the current partition
    print(f"DEBUG: Processed partition with {len(df)} rows.")
    return df

print("DEBUG: Reading dataset using Dask...")
ddf = dd.read_csv("million_rows.csv")
print("DEBUG: Starting partition processing with Dask and Numba...")

# Apply the processing function to each partition (parallelized by Dask)
processed_ddf = ddf.map_partitions(process_partition)

# Compute the result and show progress
with ProgressBar():
    result_df = processed_ddf.compute()

end_time = time.time()
print("DEBUG: Complex processing completed.")
print(f"Elapsed time for complex processing: {end_time - start_time:.6f} seconds")
print("DEBUG: Preview of processed data:")
print(result_df.head())