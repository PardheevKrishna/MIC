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
    date_end='2025-12-31'
):
    """
    1) Date generation
    2) Metric columns (Val/Var) generation
    3) DataFrame assembly
    4) Excel write
    """
    phases = ["Dates", "Metrics", "Assemble DF", "Write Excel"]
    overall = tqdm(total=len(phases), desc="Overall Progress", position=0)

    # — Phase 1: Dates —
    logging.info("Phase 1: Starting date generation…")
    start = pd.to_datetime(date_start)
    end   = pd.to_datetime(date_end)
    dates = random_dates(start, end, num_rows)
    dates_str = dates.strftime('%m/%d/%Y')
    data = {'Date': dates_str}
    logging.info("Phase 1: Date generation complete.")
    overall.update(1)

    # — Phase 2: Metrics —
    logging.info(f"Phase 2: Generating '{metric_type.upper()}' metrics…")
    for i in tqdm(range(1, 11), desc="Generating metrics", position=1):
        if metric_type.lower() == 'val':
            data[f'Val_Metric{i}'] = np.random.randint(-10**9, 10**9, size=num_rows)
            data[f'Var_Metric{i}'] = pd.NA
        elif metric_type.lower() == 'var':
            data[f'Val_Metric{i}'] = pd.NA
            data[f'Var_Metric{i}'] = np.random.randint(-10**9, 10**9, size=num_rows)
        else:
            raise ValueError("metric_type must be 'val' or 'var'")
    logging.info("Phase 2: Metric generation complete.")
    overall.update(1)

    # — Phase 3: Assemble DataFrame —
    logging.info("Phase 3: Assembling DataFrame…")
    df = pd.DataFrame(data)
    df.insert(11, '', pd.NA)  # blank column
    # enforce exact column order
    cols = (
        ['Date']
        + [f'Val_Metric{i}' for i in range(1, 11)]
        + ['']
        + [f'Var_Metric{i}' for i in range(1, 11)]
    )
    df = df[cols]
    logging.info("Phase 3: DataFrame ready.")
    overall.update(1)

    # — Phase 4: Write to Excel —
    logging.info(f"Phase 4: Writing {num_rows:,} rows to '{output_file}'…")
    df.to_excel(output_file, index=False)
    logging.info("Phase 4: Excel file written.")
    overall.update(1)
    overall.close()

    print(f"\n✓ All done – '{output_file}' is ready.")

if __name__ == '__main__':
    # Example: only Val_Metric populated
    # generate_metrics_excel(num_rows=1_000_000, metric_type='val',  output_file='val_metrics.xlsx')
    #
    # Example: only Var_Metric populated
    # generate_metrics_excel(num_rows=1_000_000, metric_type='var',  output_file='var_metrics.xlsx')
    #
    # Default call (1M rows, Val_Metric):
    generate_metrics_excel()