# run_duckdb_debug_fixed.py

import time
import duckdb
import pandas as pd
from tqdm import tqdm

# Config
CSV_FILES = ["table1.csv", "table2.csv", "table3.csv"]
CHUNK_SIZE = 100_000

def load_csv_in_chunks(con, table_name, csv_path):
    """Load a CSV into DuckDB in chunks, with tqdm and per-chunk inserts."""
    print(f"\nStarting load of {table_name} from {csv_path}")
    start_table = time.time()
    first = True
    reader = pd.read_csv(csv_path, chunksize=CHUNK_SIZE)
    for i, chunk in enumerate(tqdm(reader, desc=f"Chunks for {table_name}", unit="chunk")):
        con.register("tmp_chunk", chunk)
        if first:
            con.execute(f"CREATE OR REPLACE TABLE {table_name} AS SELECT * FROM tmp_chunk")
            first = False
            print(f"  → Created {table_name}, chunk {i+1}")
        else:
            con.execute(f"INSERT INTO {table_name} SELECT * FROM tmp_chunk")
            print(f"  → Appended chunk {i+1}")
        con.unregister("tmp_chunk")
    elapsed = time.time() - start_table
    print(f"Finished {table_name} in {elapsed:.2f} s")
    return elapsed

def main():
    print("=== DuckDB Debug Runner (Fixed SQL) ===")
    # 1. Connect to DuckDB
    print("\n[1] Connecting to DuckDB…")
    con = duckdb.connect()
    print("Connected!")

    # 2. Load each CSV with debugging
    load_times = {}
    for csv in CSV_FILES:
        table = csv.replace(".csv", "")
        t = load_csv_in_chunks(con, table, csv)
        load_times[table] = t

    total_load = sum(load_times.values())
    print("\n=== Load Summary ===")
    for tbl, tm in load_times.items():
        print(f"  {tbl}: {tm:.2f} s")
    print(f"TOTAL CSV load time: {total_load:.2f} s")

    # 3. Run the complex SQL in two stages to avoid grouping on an aggregate
    duck_sql = """
    WITH base AS (
      SELECT 
        a.key                                     AS key,
        COUNT(DISTINCT b.col1)                    AS cnt_col1,
        SUM(c.col2)                               AS sum_col2,
        AVG(a.col3)                               AS avg_col3,
        STDDEV_POP(c.col5)                        AS stddev_col5,
        COUNT(*)                                 AS total_rows
      FROM table1 AS a
      INNER JOIN table2 AS b ON a.key = b.key
      LEFT JOIN table3  AS c ON a.key = c.key
      WHERE a.col4 BETWEEN 10 AND 90
      GROUP BY a.key
      HAVING total_rows > 100
    )
    SELECT
      key,
      cnt_col1,
      sum_col2,
      CASE WHEN avg_col3 > 50 THEN 'High' ELSE 'Low' END AS Category,
      stddev_col5
    FROM base
    ORDER BY sum_col2 DESC
    LIMIT 20;
    """

    print("\n[2] Starting DuckDB query execution…")
    start_query = time.time()
    df = con.execute(duck_sql).df()
    query_time = time.time() - start_query
    print(f"Query completed in {query_time:.2f} s")

    print("\n=== Overall Summary ===")
    print(f"Total load time:      {total_load:.2f} s")
    print(f"Query time:           {query_time:.2f} s")
    print(f"Total Python time:    {total_load + query_time:.2f} s")

    print("\n[3] Preview results:")
    print(df.head().to_string(index=False))

if __name__ == "__main__":
    main()