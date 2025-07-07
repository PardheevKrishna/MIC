import pandas as pd
from openpyxl import load_workbook

def clean_excel_to_dataframe(file_path):
    # Load workbook and active sheet
    wb = load_workbook(file_path)
    ws = wb.active

    # Fill merged cell ranges with the top-left value
    for merged_range in ws.merged_cells.ranges:
        min_row, min_col, max_row, max_col = merged_range.bounds
        value = ws.cell(row=min_row, column=min_col).value
        for row in range(min_row, max_row + 1):
            for col in range(min_col, max_col + 1):
                ws.cell(row=row, column=col).value = value

    # Convert worksheet to 2D list
    data = [[cell.value for cell in row] for row in ws.iter_rows()]

    # Convert to DataFrame
    df = pd.DataFrame(data)

    # Drop empty rows and columns
    df.dropna(how='all', inplace=True)
    df.dropna(axis=1, how='all', inplace=True)
    df.reset_index(drop=True, inplace=True)

    # Heuristically find the header row (row with most non-null values)
    header_idx = df.notnull().sum(axis=1).idxmax()
    df.columns = df.iloc[header_idx]
    df = df[header_idx + 1:].reset_index(drop=True)

    # Optional: Convert to best possible dtypes
    df = df.convert_dtypes()

    return df

# Usage example:
df_clean = clean_excel_to_dataframe('your_file.xlsx')
print(df_clean.head())