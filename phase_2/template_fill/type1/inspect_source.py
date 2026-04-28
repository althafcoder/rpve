import pandas as pd
df = pd.read_excel('source.xlsx', sheet_name='Employee Details')
print('Source columns:', list(df.columns))
print('\nDataFrame shape:', df.shape)
print('\nFirst 5 rows:')
print(df.head().to_string())
print('\nSample data types:')
print(df.dtypes)
