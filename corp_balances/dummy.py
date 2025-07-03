import numpy as np
import pandas as pd
from tqdm import tqdm
from datetime import datetime, timedelta

def random_dates(start_date, end_date, n):
    """Generate n random dates between start_date and end_date."""
    start_u = start_date.toordinal()
    end_u   = end_date.toordinal()
    # random integers in [start_u, end_u]
    ordinals = np.random.randint(start_u, end_u + 1, size=n)
    return pd.to_datetime(ordinals, origin='1970-01-01', unit='D')

def generate_metrics_excel(
    num_rows=1_000_000,
    metric_type='val',            # 'val' or 'var'
    output_file='metrics.xlsx',
    date_start='2000-01-01',
    date_end='2025-12-31'
):
    """
    Generates an Excel file with:
      - Column A: 'Date' of random dates between date_start and date_end
      - Columns Val_Metric1..Val_Metric10 and Var_Metric1..Var_Metric10
        * Only the chosen metric_type is populated with random +/- million–billion values
        * The other set remains empty (but columns still present)
      - One blank column between Val_… and Var_… blocks
      - A progress bar showing column‐by‐column generation via tqdm
    """
    # parse dates
    start = pd.to_datetime(date_start)
    end   = pd.to_datetime(date_end)
    
    # prepare container for data
    data = {}
    data['Date'] = random_dates(start, end, num_rows)
    
    # create metric columns
    for i in tqdm(range(1, 11), desc='Generating metrics'):
        if metric_type.lower() == 'val':
            data[f'Val_Metric{i}'] = np.random.randint(-10**9, 10**9, size=num_rows)
            data[f'Var_Metric{i}'] = pd.NA
        elif metric_type.lower() == 'var':
            data[f'Val_Metric{i}'] = pd.NA
            data[f'Var_Metric{i}'] = np.random.randint(-10**9, 10**9, size=num_rows)
        else:
            raise ValueError("metric_type must be 'val' or 'var'")
    
    # assemble DataFrame
    df = pd.DataFrame(data)
    
    # insert blank column after Val_Metric10 (i.e. at position 11)
    df.insert(11, '', pd.NA)
    
    # reorder just to be safe
    cols = ['Date'] \
         + [f'Val_Metric{i}' for i in range(1,11)] \
         + [''] \
         + [f'Var_Metric{i}' for i in range(1,11)]
    df = df[cols]
    
    # write to Excel
    df.to_excel(output_file, index=False)
    print(f"✓ Done – saved {num_rows:,} rows to '{output_file}'")

if __name__ == '__main__':
    # Example runs:
    # To generate only Val_Metric columns:
    # generate_metrics_excel(num_rows=1_000_000, metric_type='val', output_file='val_metrics.xlsx')
    #
    # To generate only Var_Metric columns:
    # generate_metrics_excel(num_rows=1_000_000, metric_type='var', output_file='var_metrics.xlsx')
    
    # By default, this will make 1,000,000 rows with Val_Metric populated:
    generate_metrics_excel()