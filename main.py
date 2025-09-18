import pandas as pd

# read with first 3 rows as header rows
df = pd.read_excel("your_file.xlsx", header=[0, 1, 2])

# Show the dataframe with multi-level column headers
print(df.head())
