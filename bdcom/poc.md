# Running SAS Code from a .sas File Faster with Python – Proof-of-Concept (POC)

This document provides a step-by-step guide on how to execute SAS code stored in an external `.sas` file using the [saspy](https://github.com/sassoftware/saspy) library, and then leverage Python’s advanced data processing frameworks (such as Pandas, Dask, PySpark) to enhance performance. The guide includes installation, configuration, execution, and performance measurement steps.

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

- **SAS Environment:** Access to a licensed SAS installation or SAS University Edition.
- **Python 3.x:** A recent version of Python must be installed.
- **Required Python Libraries:** `saspy`, `pandas`, `dask`, and `pyspark`.

---

## Installing and Configuring SASPy

### Installing SASPy

Install SASPy using pip with the following command:

```bash
pip install saspy



Configuring the SASPy Connection

SASPy connects to your SAS environment using a configuration file (commonly named sascfg_personal.py). Create or update this file in your working directory with your SAS environment specifics. For example:

SAS_config_names = ['default']

default = {
    'java': '/usr/bin/java',    # Path to your Java executable
    'iomhost': 'localhost',     # Host name for your SAS Integration server
    'iomport': 8591,            # Port number (confirm with your SAS installation)
    'authkey': 'saspykey',      # Optional authentication key
    'encoding': 'utf-8'         # Character encoding to use
}

After configuring, test the connection with this Python snippet:

import saspy
sas = saspy.SASsession(cfgname='default')
print(sas)

If a SAS session is successfully launched, your configuration is correct.

⸻

Running SAS Code from a .sas File

Instead of embedding SAS code in your Python script, store your SAS code in an external file (e.g., my_script.sas). Follow these steps:
	1.	Create the SAS File
Save the SAS code in a file called my_script.sas. For example:

/* my_script.sas */
data work.test;
    set sashelp.class;
run;


	2.	Read and Execute the SAS Code in Python
Use Python’s file I/O functions to load the SAS code and submit it via SASPy:

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


	3.	Convert SAS Data to a Pandas DataFrame
Retrieve the SAS dataset (e.g., WORK.TEST) and convert it for further processing:

# Convert the SAS dataset 'test' in library 'work' to a Pandas DataFrame
df = sas.sd2df(table='test', libref='work')
print(df.head())



⸻

Proof-of-Concept (POC)

Below is a complete Python script that demonstrates the entire process: reading SAS code from a file, executing it, converting the resulting SAS dataset to a Pandas DataFrame, and measuring performance along the way.

import saspy
import time
import pandas as pd

# 1. Start the SAS session
sas = saspy.SASsession(cfgname='default')

# 2. Read SAS code from the file 'my_script.sas'
with open('my_script.sas', 'r') as file:
    sas_code = file.read()

# 3. Submit the SAS code and measure execution time
start_time = time.time()
result = sas.submit(sas_code)
sas_elapsed = time.time() - start_time
print("SAS Code Execution Time: {:.3f} seconds".format(sas_elapsed))
print("SAS Log:\n", result['LOG'])

# 4. Convert the SAS dataset (e.g., WORK.TEST) to a Pandas DataFrame and measure conversion time
start_time = time.time()
df = sas.sd2df(table='test', libref='work')
data_conversion_elapsed = time.time() - start_time
print("Data Conversion Time: {:.3f} seconds".format(data_conversion_elapsed))
print("Data Preview:\n", df.head())

# 5. Process data using Pandas (calculating summary statistics) and measure processing time
start_time = time.time()
summary = df.describe()
pandas_processing_elapsed = time.time() - start_time
print("Pandas Computation Time: {:.3f} seconds".format(pandas_processing_elapsed))
print("Summary Statistics:\n", summary)

What This POC Demonstrates
	•	SAS Code Execution: How to load and execute SAS code stored in an external file.
	•	Data Transfer: Converting a SAS dataset into a Pandas DataFrame.
	•	Performance Measurements: Timing the SAS code execution, data conversion, and Python-based data processing.
	•	Basic Data Analysis: Running summary statistics using Pandas.

⸻

Leveraging Python Libraries for Speed

After retrieving the data in Python, you can further optimize the workflow by using high-performance data processing libraries:

Using Pandas

Pandas offers efficient, in-memory operations optimized for vectorized computations. For example:

summary = df.describe()
print(summary)

Scaling with Dask

Dask enables parallel and out-of-core computations on large datasets. Here is an example of converting a Pandas DataFrame into a Dask DataFrame:

import dask.dataframe as dd

# Convert the Pandas DataFrame to a Dask DataFrame with 4 partitions
ddf = dd.from_pandas(df, npartitions=4)
mean_age = ddf['Age'].mean().compute()
print("Mean Age:", mean_age)

Distributed Processing with PySpark

For very large datasets, PySpark can distribute processing over a cluster. For instance:

from pyspark.sql import SparkSession

# Initialize a Spark session
spark = SparkSession.builder.appName("SAS Data Processing").getOrCreate()

# Convert the Pandas DataFrame to a Spark DataFrame
spark_df = spark.createDataFrame(df)
spark_df.createOrReplaceTempView("test")
result_df = spark.sql("SELECT AVG(Age) as avg_age FROM test")
result_df.show()

Other Techniques
	•	NumPy: Utilize for fast numerical computations.
	•	CuDF & RAPIDS: For GPU-accelerated DataFrame operations (with the appropriate hardware).
	•	Joblib/Multiprocessing: Use to parallelize independent tasks across multiple CPU cores.
	•	Cython/Numba: Optimize critical code sections by compiling to C-level speeds.

⸻

Performance Considerations and Expected Gains
	•	Reduced Overhead: Python libraries often bypass additional I/O and formatting overhead inherent in native SAS operations.
	•	Vectorized Operations: Libraries like Pandas can achieve speed-ups of 2x to 5x for moderate datasets.
	•	Parallel/Distributed Processing: Dask and PySpark can lead to improvements of 10x or more when working with large-scale datasets.
	•	Benchmarking: It is crucial to run performance tests on a representative sample of your workload to measure real-world improvements.

⸻

Conclusion

By storing SAS code in an external .sas file and executing it using SASPy, you can seamlessly integrate your SAS workflows into Python. The subsequent use of high-performance libraries like Pandas, Dask, and PySpark not only streamlines data processing but also offers significant gains in execution speed and scalability.

Happy coding and efficient data processing!

