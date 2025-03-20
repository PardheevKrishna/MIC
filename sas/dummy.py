import pandas as pd
import numpy as np
from tqdm import tqdm

# Prepare an empty list to hold rows
data = []

# Loop to generate one million rows with progress indication
for i in tqdm(range(1_000_000), desc="Generating rows"):
    row = {
        'id': i,
        'value1': np.random.random(),         # Random float between 0 and 1
        'value2': np.random.randint(0, 100)      # Random integer from 0 to 99
    }
    data.append(row)

# Convert list of dictionaries into a DataFrame
df = pd.DataFrame(data)

# Save to CSV for later use
df.to_csv("million_rows.csv", index=False)
print("Dataset saved to 'million_rows.csv'")