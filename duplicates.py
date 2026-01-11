import pandas as pd

df = pd.read_excel("HR_Analysis_Output.xlsx", sheet_name=None)

# Show all sheet names
print(df.keys())

# Show columns for the Promotion & Transfer sheet
print(df["Promotion & Transfer"].columns)
