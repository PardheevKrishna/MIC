import time
import saspy

# ===============================
# Part 1: SAS Code via saspy
# ===============================

# Create a SAS session (ensure that your SAS configuration is correct)
sas = saspy.SASsession()

print("Running SAS code via saspy...")

# Start the timer for SAS processing
start_time_sas = time.time()

# Sample SAS code: compute summary statistics on the sashelp.class dataset
sas_code = """
proc means data=sashelp.class;
   var Age Height Weight;
run;
"""

# Submit the SAS code and capture the output (this sends the code to the SAS engine)
sas_result = sas.submit(sas_code)

# Calculate elapsed time for SAS processing
sas_elapsed = time.time() - start_time_sas
print("SAS processing time: {:.2f} seconds".format(sas_elapsed))

# Optionally, print a summary of the SAS output (logs, results, etc.)
print("SAS Output Summary:")
print(sas_result['LOG'][:500])  # printing first 500 characters of the log for brevity

# ===============================
# Part 2: PySpark Code
# ===============================

from pyspark.sql import SparkSession

# Initialize a SparkSession (ensure that your Spark environment is set up)
spark = SparkSession.builder.appName("SAS_vs_Spark_Comparison").getOrCreate()

print("\nRunning PySpark code...")

# Start the timer for PySpark processing
start_time_spark = time.time()

# For comparison, create a DataFrame similar to sashelp.class.
# (In a real-world scenario, you might load the same dataset from a CSV, database, etc.)
data = [
    ("Alfred", 14, 69.0, 112.5),
    ("Alice", 13, 56.5, 84.0),
    ("Barbara", 13, 65.3, 98.0),
    ("Carol", 14, 62.8, 102.5),
    ("Henry", 14, 63.5, 102.5),
    ("James", 12, 57.3, 83.0),
    ("Jane", 12, 59.8, 84.5),
    ("Lily", 13, 56.0, 77.0),
    ("Robert", 12, 61.8, 87.0)
]
columns = ["Name", "Age", "Height", "Weight"]

# Create the DataFrame
df = spark.createDataFrame(data, columns)

# Compute summary statistics (similar to PROC MEANS)
df.describe().show()

# Calculate elapsed time for PySpark processing
spark_elapsed = time.time() - start_time_spark
print("PySpark processing time: {:.2f} seconds".format(spark_elapsed))

# Optionally, stop the Spark session when done
spark.stop()