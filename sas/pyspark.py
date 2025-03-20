import pandas as pd
import numpy as np
import dask.dataframe as dd
from dask.diagnostics import ProgressBar
from numba import njit
import time

# -------------------------------
# Pre‑processing: Load dataset
# -------------------------------
print("DEBUG: Loading dataset from CSV (excluded from timing)...")
ddf = dd.read_csv("million_rows.csv")
# Persist the dataset into memory to avoid including I/O overhead in processing time.
ddf = ddf.persist()
print("DEBUG: Dataset loaded. Rows available:", ddf.shape[0].compute())

# -----------------------------------------------
# Begin heavy processing (timed only)
# -----------------------------------------------
print("DEBUG: Starting heavy processing...")
start_time = time.time()

@njit
def compute_array(values1, values2):
    n = values1.shape[0]
    result = np.empty(n)
    for i in range(n):
        s = 0.0
        # Simulated heavy computation: inner loop of 100 iterations
        for j in range(100):
            s += np.sin(values1[i]) * np.cos(values2[i]) / (j + 1)
        result[i] = s
    return result

def process_partition(df):
    arr = compute_array(df['value1'].values, df['value2'].values)
    df['computed'] = arr
    print(f"DEBUG: Processed partition with {len(df)} rows.")
    return df

# Apply the processing function to each partition (parallelized by Dask)
processed_ddf = ddf.map_partitions(process_partition)

# Compute the result with progress monitoring
with ProgressBar():
    result_df = processed_ddf.compute()

end_time = time.time()
print("DEBUG: Heavy processing completed.")
print(f"Elapsed time for heavy processing: {end_time - start_time:.6f} seconds")
print("DEBUG: Preview of processed data:")
print(result_df.head())