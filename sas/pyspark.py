import time
import math
import numpy as np
import pandas as pd
from pyspark.sql import SparkSession
from pyspark.sql.functions import pandas_udf
from pyspark.sql.types import DoubleType

# Initialize Spark session
spark = SparkSession.builder.appName("ComplexCalculations_PandasUDF").getOrCreate()

# Generate the dummy dataset with 1,000,000 rows and 200 random columns.
df = spark.range(0, 1000000)
for i in range(1, 201):
    df = df.withColumn(f"col_{i}", np.random.rand())

# Define a vectorized function that will be applied to batches (as a Pandas DataFrame).
def complex_calculation_vectorized(pdf: pd.DataFrame) -> pd.Series:
    # Sum over the 200 columns after applying the complex calculation to each
    # We assume columns are named col_1, col_2, ..., col_200
    total = pd.Series(0.0, index=pdf.index)
    # For each column, perform the 50 iterations on the entire column at once using numpy.
    for col in pdf.columns:
        temp = pdf[col].to_numpy()
        for _ in range(50):
            temp = np.sin(temp) * np.cos(temp) + np.log(np.abs(temp) + 1)
        total += temp
    return total

# Create a Pandas UDF for vectorized operations.
complex_udf = pandas_udf(complex_calculation_vectorized, returnType=DoubleType())

# List of column names corresponding to col_1, col_2, …, col_200
columns = [f"col_{i}" for i in range(1, 201)]

# Record the start time
start_time = time.time()

# Apply the Pandas UDF on a struct of all columns.
df_complex = df.withColumn("complex_calc", complex_udf(*columns))

# Force evaluation by performing an action (e.g., counting rows)
row_count = df_complex.select("complex_calc").count()

# Record the end time
end_time = time.time()

print("Total PySpark (Pandas UDF) computation time: {:.2f} seconds".format(end_time - start_time))
df_complex.select("complex_calc").show(5)