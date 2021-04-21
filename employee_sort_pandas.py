import pandas as pd

# Combine all sheets into single dataframe
all_dfs = pd.read_excel("Employees.xlsx", sheet_name=None)
df = pd.concat(all_dfs, ignore_index=True)

# Sort the dataframe with date and turn everything into string
sorted = df.sort_values("hire_date")
sorted = sorted.astype(str)
sorted.to_excel("Combined.xlsx", index=False)