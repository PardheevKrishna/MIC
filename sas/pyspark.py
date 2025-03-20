import pandas as pd
import numpy as np
import dask.dataframe as dd
from dask.diagnostics import ProgressBar
from numba import njit

# Define a heavy computation function that operates on numpy arrays
@njit
def compute_array(values1, values2):
    n = values1.shape[0]
    result = np.empty(n)
    for i in range(n):
        s = 0.0
        # Simulate heavy computation by looping 100 times per row
        for j in range(100):
            s += np.sin(values1[i]) * np.cos(values2[i]) / (j + 1)
        result[i] = s
    return result

# Function to process each Dask partition using the numba-accelerated function
def process_partition(df):
    # Use the numba function on the partition’s numpy arrays
    arr = compute_array(df['value1'].values, df['value2'].values)
    df['computed'] = arr
    return df

# Read the pre-generated CSV using Dask (this creates a Dask DataFrame)
ddf = dd.read_csv("million_rows.csv")

# Map the processing function to each partition; Dask handles parallel processing.
processed_ddf = ddf.map_partitions(process_partition)

# Use Dask’s progress bar to monitor computation
with ProgressBar():
    result_df = processed_ddf.compute()

# Display the first few rows of the processed DataFrame
print(result_df.head())