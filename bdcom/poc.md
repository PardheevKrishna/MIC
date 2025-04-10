# Running SAS Code from a .sas File Faster with Python – All-in-One POC

This document demonstrates a complete Proof-of-Concept (POC) that reads SAS code stored in an external `.sas` file, executes it via SASPy, and then leverages multiple Python libraries—Pandas, Dask, and PySpark—for high-performance data processing. The single Python script below measures the performance at each stage to help you evaluate improvements throughout your processing pipeline.

---

## Table of Contents

- [Prerequisites](#prerequisites)
- [Installing and Configuring SASPy](#installing-and-configuring-saspy)
  - [Installing SASPy](#installing-saspy)
  - [Configuring the SASPy Connection](#configuring-the-saspy-connection)
- [Proof-of-Concept (POC): All-in-One Processing](#proof-of-concept-poc-all-in-one-processing)
- [Conclusion](#conclusion)

---

## Prerequisites

- **SAS Environment:** A licensed SAS installation or SAS University Edition.
- **Python 3.x:** Ensure a recent version is installed.
- **Required Python Libraries:** `saspy`, `pandas`, `dask`, and `pyspark` must be installed.

---

## Installing and Configuring SASPy

### Installing SASPy

Install SASPy using pip:

```bash
pip install saspy

Configuring the SASPy Connection

Create or update your SASPy configuration file (typically sascfg_personal.py) in your working directory. An example configuration for a local SAS environment is:

SAS_config_names = ['default']

default = {
    'java': '/usr/bin/java',    # Path to your Java executable
    'iomhost': 'localhost',     # SAS Integration host
    'iomport': 8591,            # Port number (verify with your SAS installation)
    'authkey': 'saspykey',      # Optional authentication key
    'encoding': 'utf-8'         # Character encoding
}

Test your configuration by launching a SAS session:

import saspy
sas = saspy.SASsession(cfgname='default')
print(sas)

If the session starts successfully, your configuration is valid.

⸻

Proof-of-Concept (POC): All-in-One Processing

Below is a complete Python script that demonstrates the entire workflow. This single code block uses SASPy to execute SAS code from a file, converts the resulting SAS dataset into a Pandas DataFrame, and then processes the data using Dask and PySpark—all while measuring performance at each step.

Assumption: A SAS script named my_script.sas exists in the working directory with content such as:

/* my_script.sas */
data work.test;
    set sashelp.class;
run;

# The following is a complete Python script demonstrating the integrated workflow.

import saspy
import time
import pandas as pd
import dask.dataframe as dd
from pyspark.sql import SparkSession

# -----------------------------------------------
# 1. Start the SAS session
# -----------------------------------------------
sas = saspy.SASsession(cfgname='default')

# -----------------------------------------------
# 2. Read SAS code from 'my_script.sas'
# -----------------------------------------------
with open('my_script.sas', 'r') as file:
    sas_code = file.read()

# -----------------------------------------------
# 3. Submit SAS code and measure execution time
# -----------------------------------------------
start_time = time.time()
result = sas.submit(sas_code)
sas_elapsed = time.time() - start_time
print("SAS Code Execution Time: {:.3f} seconds".format(sas_elapsed))
print("SAS Log:\n", result['LOG'])

# -----------------------------------------------
# 4. Convert SAS dataset (WORK.TEST) to a Pandas DataFrame
# -----------------------------------------------
start_time = time.time()
df_pandas = sas.sd2df(table='test', libref='work')
pandas_conversion_time = time.time() - start_time
print("Pandas Conversion Time: {:.3f} seconds".format(pandas_conversion_time))
print("Pandas DataFrame Preview:\n", df_pandas.head())

# -----------------------------------------------
# 5. Process data using Pandas (summary statistics)
# -----------------------------------------------
start_time = time.time()
summary = df_pandas.describe()
pandas_processing_time = time.time() - start_time
print("Pandas Processing Time: {:.3f} seconds".format(pandas_processing_time))
print("Pandas Summary Statistics:\n", summary)

# -----------------------------------------------
# 6. Convert to a Dask DataFrame and compute (e.g., mean of 'Age')
# -----------------------------------------------
start_time = time.time()
ddf = dd.from_pandas(df_pandas, npartitions=4)
mean_age_dask = ddf['Age'].mean().compute()
dask_time = time.time() - start_time
print("Dask Processing Time: {:.3f} seconds".format(dask_time))
print("Dask Mean Age: ", mean_age_dask)

# -----------------------------------------------
# 7. Initialize Spark, convert to a PySpark DataFrame, and perform an SQL query
# -----------------------------------------------
spark = SparkSession.builder.appName("SAS Data Processing").getOrCreate()
start_time = time.time()
spark_df = spark.createDataFrame(df_pandas)
spark_df.createOrReplaceTempView("test")
spark_result = spark.sql("SELECT AVG(Age) as avg_age FROM test").collect()
spark_time = time.time() - start_time
print("PySpark Processing Time: {:.3f} seconds".format(spark_time))
print("PySpark Avg Age: ", spark_result[0]['avg_age'])



⸻

Conclusion

This all-in-one POC illustrates how you can seamlessly integrate SAS code execution using SASPy with advanced data processing libraries in Python. By converting the output to a Pandas DataFrame and further leveraging Dask and PySpark for scalable computations, you can build an efficient, high-performance data processing pipeline.

Happy coding and efficient data processing!

