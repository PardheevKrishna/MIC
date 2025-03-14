import time
import logging
import pyreadstat
from tqdm import tqdm
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
# STEP 0: SETUP & GLOBALS
###############################################################################
XPT_FILE_PATH = 'dummy_data.xpt'  # Adjust if needed

def read_xpt_file():
    """
    Reads the SAS XPORT file (dummy_data.xpt) into a Pandas DataFrame.
    """
    logging.debug(f"Reading XPT file from: {XPT_FILE_PATH}")
    df, meta = pyreadstat.read_xport(XPT_FILE_PATH)
    logging.debug(f"XPT file loaded. DataFrame shape: {df.shape}")
    return df

def create_spark_df(pandas_df):
    """
    Creates a SparkSession and converts the Pandas DataFrame to a Spark DataFrame.
    """
    logging.debug("Starting SparkSession...")
    spark = SparkSession.builder.appName("ComplexTasksWithTQDM").getOrCreate()
    logging.debug("SparkSession started. Converting Pandas DataFrame to Spark DataFrame...")
    spark_df = spark.createDataFrame(pandas_df)
    logging.debug("Conversion to Spark DataFrame complete.")
    # Optional: Check row/column count (can be expensive for large data)
    row_count = spark_df.count()
    col_count = len(spark_df.columns)
    logging.debug(f"Spark DataFrame row count: {row_count}, column count: {col_count}")
    return spark, spark_df

def task_group_aggregation(spark_df):
    """
    Task 1: Group Aggregation - Compute average of num_1..num_100 grouped by cat_1.
    """
    logging.debug("Task 1: Group Aggregation started.")
    agg_exprs = [avg(col(f"num_{i}")).alias(f"avg_num_{i}") for i in range(1, 101)]
    group_means = spark_df.groupBy("cat_1").agg(*agg_exprs)
    logging.debug("Task 1 complete. (group_means DataFrame created)")
    return group_means

def task_join(spark_df, group_means):
    """
    Task 2: Join - Join avg_num_1 back to the main DataFrame.
    """
    logging.debug("Task 2: Join started.")
    # Select only cat_1 and avg_num_1 from group_means to reduce data volume
    subset = group_means.select("cat_1", "avg_num_1")
    joined_df = spark_df.join(subset, on="cat_1", how="left")
    logging.debug("Task 2 complete. (joined_df created)")
    return joined_df

def task_transform(joined_df):
    """
    Task 3: Data Transformation - Create new columns doubling num_1..num_10.
    We use tqdm to show progress for each column.
    """
    logging.debug("Task 3: Data Transformation started.")
    transformed_df = joined_df
    for i in tqdm(range(1, 11), desc="Doubling columns"):
        col_name = f"num_{i}"
        new_col = f"double_num_{i}"
        transformed_df = transformed_df.withColumn(new_col, col(col_name) * 2)
    logging.debug("Task 3 complete. (transformed_df created)")
    return transformed_df

def task_sort(transformed_df):
    """
    Task 4: Sorting - Sort by num_1.
    """
    logging.debug("Task 4: Sorting started.")
    sorted_df = transformed_df.orderBy("num_1")
    logging.debug("Task 4 complete. (sorted_df created)")
    return sorted_df

def task_pivot(spark_df):
    """
    Task 5: Pivoting - Example pivot of num_1 by cat_1.
    """
    logging.debug("Task 5: Pivoting started.")
    pivot_df = spark_df.groupBy("cat_1").pivot("cat_1").agg(avg("num_1"))
    logging.debug("Task 5 complete. (pivot_df created)")
    return pivot_df

def main():
    # Start timing
    script_start = time.time()

    # Read XPT file
    pandas_df = read_xpt_file()

    # Create Spark DF
    spark, spark_df = create_spark_df(pandas_df)

    # Define our tasks in a list so we can iterate with tqdm
    tasks = [
        ("Task 1: Group Aggregation", lambda: task_group_aggregation(spark_df)),
        ("Task 2: Join", None),            # We'll fill in once we have the result from Task 1
        ("Task 3: Transform", None),
        ("Task 4: Sort", None),
        ("Task 5: Pivot", lambda: task_pivot(spark_df))
    ]

    # We'll store intermediate results in a dict to pass between tasks
    results = {}

    # Run Task 1 separately, store result
    name, func = tasks[0]
    logging.debug(f"Starting {name}")
    group_means = func()  # group_aggregation(spark_df)
    results["group_means"] = group_means
    logging.debug(f"Completed {name}")

    # Now we update the lambda for Task 2 with the actual function call
    tasks[1] = ("Task 2: Join", lambda: task_join(spark_df, results["group_means"]))
    # Then tasks for transform and sort
    tasks[2] = ("Task 3: Transform", lambda: task_transform(results["task2"]))
    tasks[3] = ("Task 4: Sort", lambda: task_sort(results["task3"]))

    # Iterate over tasks 2..5 with a tqdm progress bar
    for (task_name, task_func) in tqdm(tasks[1:], desc="Main Steps"):
        logging.debug(f"Starting {task_name}")
        step_result = task_func()
        # Save step result in results dict with a short key
        short_key = task_name.lower().split()[1]  # e.g. "join", "transform", "sort", "pivot"
        results[f"task{short_key}"] = step_result
        logging.debug(f"Completed {task_name}")

    # Show a sample of the final sorted data
    sorted_df = results["tasksort"]
    logging.debug("Showing sample of the final sorted DataFrame (first 5 rows):")
    sorted_df.show(5)

    # Also show sample of the pivot
    pivot_df = results["taskpivot"]
    logging.debug("Showing sample of pivoted DataFrame (first 5 rows):")
    pivot_df.show(5)

    # Stop Spark
    spark.stop()

    script_end = time.time()
    logging.debug(f"Total script runtime (seconds): {script_end - script_start:.2f}")

if __name__ == "__main__":
    main()