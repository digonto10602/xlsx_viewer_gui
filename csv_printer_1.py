import pandas as pd

path = "RMG Leads USA.xlsx"

# sheet_name=5 means the 6th sheet (0-based indexing in pandas)
df = pd.read_excel(path, sheet_name=4)

# 1) DataFrame row index labels (left-side index)
print("Row index labels (df.index):")
print(df.index.tolist())

# 2) Column names (the labels across the top / x-axis in a typical plot)
print("\nColumn names (df.columns):")
print(df.columns.tolist())

A = df.columns.tolist()

for i in range(len(A)):
    print("ind names:", A[i])

B = df['Website'].to_numpy()

for i in range(len(B)):
    print("Website no. : ",i," link : ", B[i])

# 3) If you actually want the first column to become the index (common in Excel files):
# df2 = pd.read_excel(path, sheet_name=5, index_col=0)
# print("\nIndex labels after using first column as index (index_col=0):")
# print(df2.index.tolist())
# print("Column names after index_col=0:")
# print(df2.columns.tolist())