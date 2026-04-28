import pandas as pd
import openpyxl

# List all sheets
wb = openpyxl.load_workbook('source.xlsx', read_only=True, data_only=True)
print('All sheets in source.xlsx:')
print(wb.sheetnames)
wb.close()

# Try reading first sheet and print info
print('\n=== First sheet (index 0) ===')
try:
    df = pd.read_excel('source.xlsx', sheet_name=0)
    print('Columns:', list(df.columns))
    print('Shape:', df.shape)
    print('\nHead:')
    print(df.head(3).to_string())
except Exception as e:
    print('Error:', e)

# Try other possible sheet names
possible_sheets = ['Sheet1', 'Data', 'Employees', 'Sheet', 'Sheet 1']
for sheet in possible_sheets[1:]:  # Skip Sheet1 if first
    try:
        print(f'\n=== Sheet "{sheet}" ===')
        df = pd.read_excel('source.xlsx', sheet_name=sheet)
        print('Found! Columns:', list(df.columns))
        print('Shape:', df.shape)
    except:
        pass
