# Running SAS Code from a .sas File Faster with Python – Proof-of-Concept (POC)

This guide explains how to execute SAS code stored in an external `.sas` file using the [saspy](https://github.com/sassoftware/saspy) library, and then leverage Python’s data processing frameworks (e.g., Pandas, Dask, PySpark) for faster execution and further processing. At the end of this document, a proof-of-concept (POC) illustrates the complete workflow with performance measurements.

---

## Table of Contents

- [Prerequisites](#prerequisites)
- [Installing and Configuring SASPy](#installing-and-configuring-saspy)
  - [Installing SASPy](#installing-saspy)
  - [Configuring the SASPy Connection](#configuring-the-saspy-connection)
- [Running SAS Code from a .sas File](#running-sas-code-from-a-sas-file)
- [Proof-of-Concept (POC)](#proof-of-concept-poc)
- [Leveraging Python Libraries for Speed](#leveraging-python-libraries-for-speed)
  - [Using Pandas](#using-pandas)
  - [Scaling with Dask](#scaling-with-dask)
  - [Distributed Processing with PySpark](#distributed-processing-with-pyspark)
  - [Other Techniques](#other-techniques)
- [Performance Considerations and Expected Gains](#performance-considerations-and-expected-gains)
- [Conclusion](#conclusion)

---

## Prerequisites

- **SAS Environment:** A licensed SAS installation (or SAS University Edition) must be available.
- **Python 3.x:** Ensure a recent version of Python is installed.
- **Required Libraries:** Install Python packages such as `saspy`, `pandas`, `dask`, and `pyspark`.

---

## Installing and Configuring SASPy

### Installing SASPy

Install the SASPy library with pip:

```bash
pip install saspy


## Configuring the SASPy Connection

SASPy uses a configuration file (commonly sascfg_personal.py) to connect to your SAS session. Create or update the file in your working directory with an example configuration for a local SAS installation:

SAS_config_names = ['default']

default = {
    'java': '/usr/bin/java',    # Path to your Java executable
    'iomhost': 'localhost',     # Host name for your SAS Integration server
    'iomport': 8591,            # Port number (verify this with your SAS installation)
    'authkey': 'saspykey',      # Optional authentication key
    'encoding': 'utf-8'         # Encoding to use
}

Test the configuration by launching a SAS session:

import saspy
sas = saspy.SASsession(cfgname='default')
print(sas)

If the session launches successfully, your configuration is correct.

⸻

Running SAS Code from a .sas File

Instead of hardcoding SAS code within your Python script, you can store it in an external file. Follow these steps:
	1.	Create the SAS File:
Save your SAS code in a file named my_script.sas. For instance:

/* my_script.sas */
data work.test;
    set sashelp.class;
run;


	2.	Read and Execute the SAS Code in Python:
Read the SAS file content into a Python variable and submit it via SASPy:

import saspy

# Start the SAS session
sas = saspy.SASsession(cfgname='default')

# Read SAS code from the file
with open('my_script.sas', 'r') as file:
    sas_code = file.read()

# Submit the SAS code
result = sas.submit(sas_code)
print("SAS Log:")
print(result['LOG'])


	3.	Convert SAS Data to Pandas DataFrame:
Retrieve the SAS dataset (e.g., WORK.TEST) and convert it for further processing:

# Convert the SAS dataset 'test' to a Pandas DataFrame
df = sas.sd2df(table='test', libref='work')
print(df.head())



⸻

Proof-of-Concept (POC)

Below is a complete Python script that demonstrates the end-to-end process. This POC reads SAS code from my_script.sas, submits it, converts the resulting SAS dataset to a Pandas DataFrame, and performs simple computations with performance timing.

import saspy
import time
import pandas as pd

# 1. Start the SAS session
sas = saspy.SASsession(cfgname='default')

# 2. Read SAS code from the file 'my_script.sas'
with open('my_script.sas', 'r') as file:
    sas_code = file.read()

# 3. Measure and submit the SAS code execution
start_time = time.time()
result = sas.submit(sas_code)
sas_elapsed = time.time() - start_time
print("SAS Code Execution Time: {:.3f} seconds".format(sas_elapsed))
print("SAS Log:\n", result['LOG'])

# 4. Convert the SAS dataset (e.g., WORK.TEST) into a Pandas DataFrame
start_time = time.time()
df = sas.sd2df(table='test', libref='work')
data_conversion_elapsed = time.time() - start_time
print("Data Conversion Time: {:.3f} seconds".format(data_conversion_elapsed))
print("Data Preview:\n", df.head())

# 5. Process data using Pandas (calculate summary statistics)
start_time = time.time()
summary = df.describe()
pandas_processing_elapsed = time.time() - start_time
print("Pandas Computation Time: {:.3f} seconds".format(pandas_processing_elapsed))
print("Summary Statistics:\n", summary)

What This POC Demonstrates
	•	SAS Code Execution: Reading SAS code from an external file and submitting it using SASPy.
	•	Data Transfer: Converting a SAS dataset into a Pandas DataFrame for further manipulation.
	•	Performance Measurement: Timing the SAS code execution, data transfer, and subsequent Python data processing.
	•	Basic Data Analysis: Using Pandas to generate summary statistics from the dataset.

⸻

Leveraging Python Libraries for Speed

After acquiring the data in Python, you can harness other libraries to further accelerate processing:

Using Pandas
	•	Example:

summary = df.describe()
print(summary)


	•	Benefits: Fast, in-memory operations suitable for datasets that fit into RAM.

Scaling with Dask
	•	Example:

import dask.dataframe as dd

# Convert the Pandas DataFrame to a Dask DataFrame (e.g., 4 partitions)
ddf = dd.from_pandas(df, npartitions=4)
mean_age = ddf['Age'].mean().compute()
print("Mean Age:", mean_age)


	•	Benefits: Parallelized and out-of-core processing for larger-than-memory datasets.

Distributed Processing with PySpark
	•	Example:

from pyspark.sql import SparkSession

# Initialize a Spark session
spark = SparkSession.builder.appName("SAS Data Processing").getOrCreate()

# Convert the Pandas DataFrame to a Spark DataFrame
spark_df = spark.createDataFrame(df)
spark_df.createOrReplaceTempView("test")
result_df = spark.sql("SELECT AVG(Age) as avg_age FROM test")
result_df.show()


	•	Benefits: Handling very large datasets using distributed processing.

Other Techniques
	•	NumPy: Efficient numerical operations.
	•	CuDF & RAPIDS: GPU-accelerated DataFrame processing (with compatible hardware).
	•	Joblib/Multiprocessing: Parallelizing tasks on multiple CPU cores.
	•	Cython/Numba: Compiling Python code to C-level speeds for critical operations.

⸻

Performance Considerations and Expected Gains
	•	Reduced Overhead: Python often bypasses some of the I/O and formatting overhead seen with native SAS execution.
	•	Vectorized Operations: Libraries like Pandas are optimized for vectorized operations, yielding speed-ups between 2x to 5x for moderate datasets.
	•	Parallel Processing: Utilizing Dask or PySpark for distributed or out-of-core computations can lead to 10x improvements or more on large-scale data.
	•	Benchmarking: Run tests on a representative subset of your workload to measure real-world improvements.

⸻

Conclusion

By reading SAS code from an external .sas file and executing it via SASPy, you can integrate your existing SAS workflows into Python seamlessly. With the added power of Python’s data processing libraries, such as Pandas, Dask, and PySpark, you can achieve significant performance improvements and more flexible data processing pipelines.

Happy coding and efficient data processing!

