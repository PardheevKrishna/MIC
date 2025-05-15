# generate_dummy.py

import numpy as np
import pandas as pd
from tqdm import tqdm

n_rows = 10**6
n_cols = 1000
chunk_size = 10**5   # write in 10 chunks of 100k rows
col_names = [f"col{i}" for i in range(1, n_cols+1)]

for tbl in ("table1","table2","table3"):
    for i in tqdm(range(n_rows // chunk_size), desc=f"Writing {tbl}.csv"):
        data = np.random.randint(0, 100, size=(chunk_size, n_cols))
        df = pd.DataFrame(data, columns=col_names)
        df["key"] = np.random.randint(1, 200_000, size=chunk_size)
        df.to_csv(f"{tbl}.csv", mode="a", index=False, header=(i==0))