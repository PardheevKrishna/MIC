import numpy as np
import pandas as pd
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor
from tqdm import tqdm
import xlsxwriter

def random_dates(start: pd.Timestamp, end: pd.Timestamp, n: int) -> pd.DatetimeIndex:
    """Vectorized: n random dates between start/end at midnight."""
    span = (end - start).days
    offsets = np.random.randint(0, span+1, size=n)
    return start + pd.to_timedelta(offsets, unit='D')

def _gen_block(names, metric_type, n):
    """Worker: generate one block (Val_… or Var_…) in one thread."""
    block = {}
    is_val = metric_type == 'val'
    for name in names:
        if name.startswith('Val_') and is_val:
            block[name] = np.random.randint(-10**9, 10**9, size=n)
        elif name.startswith('Var_') and not is_val:
            block[name] = np.random.randint(-10**9, 10**9, size=n)
        else:
            # keep the column but empty
            block[name] = [pd.NA] * n
    return block

def generate_metrics_excel(
    num_rows: int = 1_000_000,
    metric_type: str = 'val',          # 'val' or 'var'
    output_file: str = 'metrics.xlsx',
    date_start: str = '2000-01-01',
    date_end:   str = '2025-12-31',
    row_chunk:  int = 50_000           # buffer size for row‐writes
):
    # 1) Prepare Date column
    start = pd.to_datetime(date_start)
    end   = pd.to_datetime(date_end)
    dates = random_dates(start, end, num_rows).strftime('%m/%d/%Y')
    
    # 2) Kick off two threads to build Val_Metric1–10 and Var_Metric1–10
    val_names = [f'Val_Metric{i}' for i in range(1, 11)]
    var_names = [f'Var_Metric{i}' for i in range(1, 11)]
    with ThreadPoolExecutor(max_workers=2) as ex:
        fv = ex.submit(_gen_block, val_names, metric_type.lower(), num_rows)
        fv2 = ex.submit(_gen_block, var_names, metric_type.lower(), num_rows)
    blocks = {}
    blocks.update(fv.result())
    blocks.update(fv2.result())
    
    # 3) Setup xlsxwriter in constant_memory mode
    workbook  = xlsxwriter.Workbook(output_file, {'constant_memory': True})
    worksheet = workbook.add_worksheet()
    
    # 4) Write header row
    headers = ['Date'] + val_names + [''] + var_names
    worksheet.write_row(0, 0, headers)
    
    # 5) Stream‐write each row with tqdm progress
    total = num_rows
    fmt_row = [dates] + [blocks[n] for n in val_names] + [None] + [blocks[n] for n in var_names]
    # We’ll buffer `row_chunk` rows at a time for lower method‐call overhead
    buffer = []
    row_idx = 1
    pbar = tqdm(total=total, desc="Writing rows", unit="rows")
    for i in range(total):
        # build one row as a plain Python list
        row = [
            dates[i],
            * (blocks[name][i] for name in val_names),
            None,
            * (blocks[name][i] for name in var_names)
        ]
        buffer.append(row)
        if len(buffer) >= row_chunk:
            # write out this batch
            for buf in buffer:
                worksheet.write_row(row_idx, 0, buf)
                row_idx += 1
            buffer.clear()
            pbar.update(row_chunk)
    # final flush
    for buf in buffer:
        worksheet.write_row(row_idx, 0, buf)
        row_idx += 1
    pbar.update(len(buffer))
    pbar.close()
    
    # 6) Close workbook
    workbook.close()
    print(f"\n✓ Done: {num_rows:,} rows → {output_file}")

if __name__ == '__main__':
    # Example: val‐metrics only
    generate_metrics_excel(
        num_rows=1_000_000,
        metric_type='val',
        output_file='val_metrics_fast.xlsx'
    )
    # Or for var‐metrics:
    # generate_metrics_excel(metric_type='var', output_file='var_metrics_fast.xlsx')