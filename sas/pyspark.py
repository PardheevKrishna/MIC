import time
import logging
import pyreadstat
from pyspark.sql import SparkSession
from pyspark.sql.functions import avg, col

###############################################################################
# CONFIGURE LOGGING
###############################################################################
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

###############################################################################
# START RUNTIME MEASUREMENT
###############################################################################
script_start = time.time()
logging.debug("Script started. Initializing PySpark job...")

###############################################################################
# 1. READ THE SAS XPORT FILE INTO A PANDAS DATAFRAME
###############################################################################
xpt_file_path = 'dummy_data.xpt'
logging.debug(f"Reading XPT file from: {xpt_file_path}")
df, meta = pyreadstat.read_xport(xpt_file_path)
logging.debug(f"Successfully loaded DataFrame from XPT. Shape: {df.shape}")

###############################################################################
# 2. START A SPARK SESSION AND CONVERT TO SPARK DATAFRAME
###############################################################################
spark_start = time.time()
logging.debug("Starting SparkSession...")
spark = SparkSession.builder.appName("ComplexTasksDebug").getOrCreate()
logging.debug("SparkSession started.")

logging.debug("Converting Pandas DataFrame to Spark DataFrame...")
spark_df = spark.createDataFrame(df)
logging.debug("Conversion complete. Spark DataFrame created.")

# (Optional) Count the rows for a sanity check (can be expensive for large data)
row_count = spark_df.count()
col_count = len(spark_df.columns)
logging.debug(f"Spark DataFrame row count: {row_count}, column count: {col_count}")

spark_end = time.time()
logging.debug(f"Spark session init + DataFrame creation took {spark_end - spark_start:.2f} seconds.")

###############################################################################
# 3. TASK 1: GROUP AGGREGATION (COMPUTE AVERAGE OF NUM_1 - NUM_100 BY CAT_1)
###############################################################################
task1_start = time.time()
logging.debug("Starting Task 1: Group Aggregation (num_1 to num_100) by cat_1...")

agg_exprs = [avg(col(f"num_{i}")).alias(f"avg_num_{i}") for i in range(1, 101)]
group_means = spark_df.groupBy("cat_1").agg(*agg_exprs)

logging.debug("Task 1 complete. Sample of group_means schema:")
group_means.printSchema()

task1_end = time.time()
logging.debug(f"Task 1 duration: {task1_end - task1_start:.2f} seconds.")

###############################################################################
# 4. TASK 2: JOIN THE AVERAGE OF NUM_1 BACK TO THE MAIN DATAFRAME
###############################################################################
task2_start = time.time()
logging.debug("Starting Task 2: Joining avg_num_1 back to main DataFrame...")

# Select only cat_1 and avg_num_1 from group_means to reduce data for the join
group_subset = group_means.select("cat_1", "avg_num_1")

joined_df = spark_df.join(group_subset, on="cat_1", how="left")

logging.debug("Task 2 complete. Joined DataFrame schema:")
joined_df.printSchema()

task2_end = time.time()
logging.debug(f"Task 2 duration: {task2_end - task2_start:.2f} seconds.")

###############################################################################
# 5. TASK 3: DATA TRANSFORMATION (DOUBLE num_1 - num_10)
###############################################################################
task3_start = time.time()
logging.debug("Starting Task 3: Creating new columns (double_num_1 to double_num_10)...")

transformed_df = joined_df
for i in range(1, 11):
    col_name = f"num_{i}"
    new_col = f"double_num_{i}"
    transformed_df = transformed_df.withColumn(new_col, col(col_name) * 2)

logging.debug("Task 3 complete. Sample of transformed DataFrame schema:")
transformed_df.printSchema()

task3_end = time.time()
logging.debug(f"Task 3 duration: {task3_end - task3_start:.2f} seconds.")

###############################################################################
# 6. TASK 4: SORT THE DATA BY num_1
###############################################################################
task4_start = time.time()
logging.debug("Starting Task 4: Sorting by num_1...")

sorted_df = transformed_df.orderBy("num_1")

logging.debug("Task 4 complete. Showing first 5 rows of sorted data for debug:")
sorted_df.show(5)

task4_end = time.time()
logging.debug(f"Task 4 duration: {task4_end - task4_start:.2f} seconds.")

###############################################################################
# 7. TASK 5: PIVOTING (EXAMPLE: PIVOT num_1 BY cat_1)
###############################################################################
task5_start = time.time()
logging.debug("Starting Task 5: Pivoting (e.g., average of num_1 pivoted by cat_1)...")

pivot_df = spark_df.groupBy("cat_1").pivot("cat_1").agg(avg("num_1"))

logging.debug("Task 5 complete. Pivoted DataFrame schema:")
pivot_df.printSchema()

task5_end = time.time()
logging.debug(f"Task 5 duration: {task5_end - task5_start:.2f} seconds.")

###############################################################################
# FINAL RUNTIME REPORT
###############################################################################
script_end = time.time()
total_elapsed = script_end - script_start
logging.debug(f"Total script runtime (seconds): {total_elapsed:.2f}")

logging.debug("Done. Stopping SparkSession.")
spark.stop()