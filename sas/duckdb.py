# run_duckdb_timed.py

import time
import duckdb

# 1. Connect to DuckDB in-process
con = duckdb.connect()

# 2. Load CSVs (fast, parallel) using read_csv_auto
start_load = time.time()
con.execute("CREATE OR REPLACE TABLE table1 AS SELECT * FROM read_csv_auto('table1.csv')")
con.execute("CREATE OR REPLACE TABLE table2 AS SELECT * FROM read_csv_auto('table2.csv')")
con.execute("CREATE OR REPLACE TABLE table3 AS SELECT * FROM read_csv_auto('table3.csv')")
end_load = time.time()

# 3. Define the complex SQL
duck_sql = """
SELECT 
  a.key,
  COUNT(DISTINCT b.col1)   AS cnt_col1,
  SUM(c.col2)              AS sum_col2,
  CASE WHEN AVG(a.col3)>50 THEN 'High' ELSE 'Low' END AS Category,
  STDDEV_POP(c.col5)       AS stddev_col5
FROM table1 AS a
  INNER JOIN table2 AS b ON a.key = b.key
  LEFT  JOIN table3 AS c ON a.key = c.key
WHERE a.col4 BETWEEN 10 AND 90
GROUP BY a.key, Category
HAVING COUNT(*) > 100
ORDER BY sum_col2 DESC
LIMIT 20
"""

# 4. Run and time the query
start_query = time.time()
df = con.execute(duck_sql).df()
end_query = time.time()

# 5. Report timings
print(f"CSV load time:     {end_load - start_load:.2f} s")
print(f"DuckDB query time: {end_query - start_query:.2f} s")
print(f"Total Python time: {end_query - start_load:.2f} s\n")

# 6. Show a few results
print(df.head())