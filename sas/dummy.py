import numpy as np
import pandas as pd
from tqdm import tqdm
import pyreadstat

# Parameters
n_rows = 10**6
n_cols = 1000
col_names = [f'col{i}' for i in range(n_cols)]
chunksize = 100000  # adjust based on your available RAM

# 1. Generate dummy data in chunks
dfs = []
for _ in tqdm(range(n_rows // chunksize), desc="Generating data"):
    data = np.random.rand(chunksize, n_cols)
    df_chunk = pd.DataFrame(data, columns=col_names)
    dfs.append(df_chunk)

# 2. Concatenate all chunks
df = pd.concat(dfs, ignore_index=True)

# 3. Write to SAS7BDAT
pyreadstat.write_sas7bdat(df, 'dummy_data.sas7bdat')

print("Created dummy_data.sas7bdat (1M rows × 1000 cols)")