import time
import pyreadstat
from pyspark.sql import SparkSession
from pyspark.sql.functions import avg, col

# ----------------------------------------------------------
# Python (PySpark) Code: Import dummy_data.xpt and perform complex tasks
# ----------------------------------------------------------

# Start runtime measurement
start_time = time.time()

# --- Step 1: Read the SAS XPORT file ---
# This reads the .xpt file into a Pandas DataFrame.
df, meta = pyreadstat.read_xport('dummy_data.xpt')

# --- Step 2: Start SparkSession and convert to Spark DataFrame ---
spark = SparkSession.builder.appName("ComplexTasks").getOrCreate()
spark_df = spark.createDataFrame(df)

# --- Task 1: Group Aggregation ---
# Compute the average of num_1 to num_100 grouped by cat_1.
agg_exprs = [avg(col(f"num_{i}")).alias(f"avg_num_{i}") for i in range(1, 101)]
group_means = spark_df.groupBy("cat_1").agg(*agg_exprs)

# --- Task 2: Join ---
# Join the average of num_1 back to the main DataFrame.
joined_df = spark_df.join(group_means.select("cat_1", "avg_num_1"), on="cat_1", how="left")

# --- Task 3: Data Transformation ---
# Create new columns doubling num_1 to num_10.
for i in range(1, 11):
    joined_df = joined_df.withColumn(f"double_num_{i}", col(f"num_{i}") * 2)

# --- Task 4: Sorting ---
# Sort the DataFrame by num_1.
sorted_df = joined_df.orderBy("num_1")

# --- Task 5: Pivoting ---
# Example pivot: compute the average of num_1 for each category and pivot on cat_1.
pivot_df = spark_df.groupBy("cat_1").pivot("cat_1").agg(avg("num_1"))

# Stop runtime measurement
end_time = time.time()
elapsed = end_time - start_time
print("Total PySpark runtime (seconds):", elapsed)

# Show a few rows of one of the results for verification.
sorted_df.select("cat_1", "num_1", "avg_num_1", "double_num_1").show(5)

# Stop the Spark session when done.
spark.stop()