import pandas as pd

df_dict = pd.read_excel('HR_Analysis_Output.xlsx', sheet_name=None)
print('Sheets:', list(df_dict.keys()))
for sheet, data in df_dict.items():
    print(f'{sheet} columns: {list(data.columns)}')