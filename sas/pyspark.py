import time
import math
from pyspark.sql import SparkSession
from pyspark.sql.functions import rand, udf
from pyspark.sql.types import DoubleType

# Initialize the Spark session
spark = SparkSession.builder.appName("ComplexCalculations").getOrCreate()

# Generate the dummy dataset with 1,000,000 rows and 200 random columns.
df = spark.range(0, 1000000)
for i in range(1, 201):
    df = df.withColumn(f"col_{i}", rand())

# Define the UDF for the complex calculation
def complex_calculation(*cols):
    total = 0.0
    for x in cols:
        temp = x
        for _ in range(50):
            temp = math.sin(temp) * math.cos(temp) + math.log(abs(temp) + 1)
        total += temp
    return total

complex_udf = udf(complex_calculation, DoubleType())

# List of column names corresponding to col_1, col_2, …, col_200
columns = [f"col_{i}" for i in range(1, 201)]

# Record the start time
start_time = time.time()

# Apply the UDF to compute a new column "complex_calc"
df_complex = df.withColumn("complex_calc", complex_udf(*columns))

# Force evaluation by performing an action (here, counting rows)
df_complex.select("complex_calc").count()

# Record the end time
end_time = time.time()

print("Total PySpark computation time: {:.2f} seconds".format(end_time - start_time))