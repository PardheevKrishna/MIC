import numpy as np
import pandas as pd
from tqdm import tqdm
from datetime import datetime
import logging

# ——— configure logging ———
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s %(levelname)s: %(message)s',
    datefmt='%H:%M:%S'
)

def random_dates(start_date, end_date, n):
    """Generate n random dates between start_date and end_date."""
    delta_days = (end_date - start_date).days
    offsets = np.random.randint(0, delta_days + 1, size=n)
    return start_date + pd.to_timedelta(offsets, unit='D')

def generate_metrics_excel(
    num_rows=1_000_000,
    metric_type='val',            # 'val' or 'var'
    output_file='metrics.xlsx',
    date_start='2000-01-01',
    date_end='2025-12-31',
    chunk_size=1_000              # rows per write-chunk for progress
):
    """
    1) Date generation
    2) Metric columns generation
    3) DataFrame assembly
    4) Writing to Excel in chunks with tqdm
    """
    phases = ["Dates", "Metrics", "Assemble DF", "Write Excel"]
    overall = tqdm(total=len(phases), desc="Overall Progress", position=0)

    # Phase 1: Dates
    logging.info("Phase 1: Starting date generation…")
    start = pd.to_datetime(date_start)
    end   = pd.to_datetime(date_end)
    dates = random_dates(start, end, num_rows)
    data = {'Date': dates.strftime('%m/%d/%Y')}
    logging.info("Phase 1: Date generation complete.")
    overall.update(1)

    # Phase 2: Metrics
    logging.info(f"Phase 2: Generating '{metric_type.upper()}' metrics…")
    for i in tqdm(range(1, 11), desc="Generating metrics", position=1):
        if metric_type.lower() == 'val':
            data[f'Val_Metric{i}'] = np.random.randint(-10**9, 10**9, size=num_rows)
            data[f'Var_Metric{i}'] = pd.NA
        else:
            data[f'Val_Metric{i}'] = pd.NA
            data[f'Var_Metric{i}'] = np.random.randint(-10**9, 10**9, size=num_rows)
    logging.info("Phase 2: Metric generation complete.")
    overall.update(1)

    # Phase 3: Assemble DataFrame
    logging.info("Phase 3: Assembling DataFrame…")
    df = pd.DataFrame(data)
    df.insert(11, '', pd.NA)  # blank column
    cols = (
        ['Date']
        + [f'Val_Metric{i}' for i in range(1, 11)]
        + ['']
        + [f'Var_Metric{i}' for i in range(1, 11)]
    )
    df = df[cols]
    logging.info("Phase 3: DataFrame ready.")
    overall.update(1)

    # Phase 4: Write to Excel with per-chunk progress
    logging.info(f"Phase 4: Writing {num_rows:,} rows to '{output_file}' in chunks of {chunk_size}…")
    overall.update(1)
    overall.close()

    # Use xlsxwriter so we can specify startrow
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        # write header
        df.iloc[0:0].to_excel(writer, index=False)
        worksheet = writer.sheets['Sheet1']

        # iterate chunks
        for start_row in tqdm(range(0, num_rows, chunk_size),
                              desc="Writing to Excel",
                              position=0):
            end_row = min(start_row + chunk_size, num_rows)
            df_chunk = df.iloc[start_row:end_row]
            # +1 because header is row 0
            df_chunk.to_excel(
                writer,
                index=False,
                header=False,
                startrow=start_row + 1,
                sheet_name='Sheet1'
            )
        writer.save()
    logging.info("Phase 4: Excel file written.")
    print(f"\n✓ All done – '{output_file}' is ready.")

if __name__ == '__main__':
    # Default call:
    generate_metrics_excel()