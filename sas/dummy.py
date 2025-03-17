import pandas as pd
import numpy as np
import xlsxwriter

# -----------------------------
# Step 1: Set File Name & Parameters
# -----------------------------
output_file = "dummy_data.xlsx"
n_rows = 1_000_000  # 1 million rows
n_cols = 200        # 200 columns

# Create column names: var1, var2, ..., var200
columns = [f'var{i}' for i in range(1, n_cols + 1)]

# -----------------------------
# Step 2: Create an Excel File with xlsxwriter (Row-by-Row)
# -----------------------------
# Initialize the workbook and worksheet
workbook = xlsxwriter.Workbook(output_file)
worksheet = workbook.add_worksheet("Sheet1")

# Write the header row
worksheet.write_row(0, 0, columns)

# -----------------------------
# Step 3: Write Data Row-by-Row with Progress
# -----------------------------
print(f"Writing {n_rows} rows to {output_file}...")

for row_num in range(1, n_rows + 1):
    # Generate a row of random data
    row_data = np.random.randn(n_cols).tolist()

    # Write row to the Excel file
    worksheet.write_row(row_num, 0, row_data)

    # Print progress every 10,000 rows
    if row_num % 10000 == 0:
        print(f"{row_num}/{n_rows} rows written...")

# Close the workbook (finalize the file)
workbook.close()

print(f"Excel file successfully created: {output_file}")