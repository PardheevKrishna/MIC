import numpy as np
import pandas as pd
from tqdm import tqdm
from datetime import datetime
from datetime import timedelta

def random_dates(start_date, end_date, n):
    """
    Generate n random dates between start_date and end_date.
    Returns a DatetimeIndex at midnight for each date.
    """
    delta_days = (end_date - start_date).days
    # random offsets from 0 to delta_days
    random_offsets = np.random.randint(0, delta_days + 1, size=n)
    return start_date + pd.to_timedelta(random_offsets, unit='D')

def generate_metrics_excel(
    num_rows=1_000_000,
    metric_type='val',            # 'val' or 'var'
    output_file='metrics.xlsx',
    date_start='2000-01-01',
    date_end='2025-12-31'
):
    """
    - Column A: 'Date' in MM/DD/YYYY format (no time)
    - Columns Val_Metric1..Val_Metric10 and Var_Metric1..Var_Metric10
      * Only the chosen metric_type gets random +/- values up to ±1e9
      * The other block is left empty
    - One blank column between the two metric blocks
    - Progress bar via tqdm
    """
    # parse string dates to Timestamps (at midnight)
    start = pd.to_datetime(date_start)
    end   = pd.to_datetime(date_end)
    
    # build data dict
    data = {}
    dates = random_dates(start, end, num_rows)
    # format as MM/DD/YYYY strings
    data['Date'] = dates.strftime('%m/%d/%Y')
    
    # generate your 10 val & 10 var columns
    for i in tqdm(range(1, 11), desc='Generating metrics'):
        if metric_type.lower() == 'val':
            data[f'Val_Metric{i}'] = np.random.randint(-10**9, 10**9, size=num_rows)
            data[f'Var_Metric{i}'] = pd.NA
        elif metric_type.lower() == 'var':
            data[f'Val_Metric{i}'] = pd.NA
            data[f'Var_Metric{i}'] = np.random.randint(-10**9, 10**9, size=num_rows)
        else:
            raise ValueError("metric_type must be 'val' or 'var'")
    
    # assemble into DataFrame
    df = pd.DataFrame(data)
    # insert one blank column after Val_Metric10 (position index 11)
    df.insert(11, '', pd.NA)
    
    # enforce the desired column order
    cols = ['Date'] \
         + [f'Val_Metric{i}' for i in range(1,11)] \
         + [''] \
         + [f'Var_Metric{i}' for i in range(1,11)]
    df = df[cols]
    
    # write out to Excel
    df.to_excel(output_file, index=False)
    print(f"✓ Done – {num_rows:,} rows written to '{output_file}'")

if __name__ == '__main__':
    # Example: only Val_Metric populated
    # generate_metrics_excel(num_rows=1_000_000, metric_type='val',  output_file='val_metrics.xlsx')
    #
    # Example: only Var_Metric populated
    # generate_metrics_excel(num_rows=1_000_000, metric_type='var',  output_file='var_metrics.xlsx')
    #
    # Default call (1M rows, Val_Metric):
    generate_metrics_excel()