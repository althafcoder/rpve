import pandas as pd
import openpyxl

print('Inspecting source.xlsx for plan name and current premium...')

sheets = pd.ExcelFile('source.xlsx').sheet_names
print('Sheets:', sheets)

for sheet in ['Census', 'SalesConsultantBroker List', 'Plan Network Test', 'Region Selection']:
    try:
        print(f'\n=== {sheet} ===')
        # Try different skiprows for headers
        for skip in [0,1,2,3,4,5]:
            df = pd.read_excel('source.xlsx', sheet_name=sheet, skiprows=skip, nrows=3)
            if not df.empty and len(df.columns) > 0:
                print(f'Skiprows={skip} columns sample: {list(df.columns[:5])}')
                print(df.iloc[0].to_string()[:200] + '...' if len(df.iloc[0].to_string()) > 200 else df.iloc[0].to_string())
                # Look for keywords
                cols_str = ' '.join([str(c) for c in df.columns]).lower()
                if 'plan' in cols_str or 'premium' in cols_str or 'current' in cols_str:
                    print('*** KEYWORDS FOUND ***')
        break
    except Exception as e:
        print(f'Error: {e}')

print('\nRaw Census rows 0-10 columns 0-20:')
df = pd.read_excel('source.xlsx', sheet_name='Census', nrows=10)
print(df.iloc[:, :20].to_string())

print('\\nSearch for headers:')
wb = openpyxl.load_workbook('source.xlsx', data_only=True)
ws = wb['Census']
for row in ws.iter_rows(min_row=1, max_row=10, values_only=True):
    row_str = ' | '.join([str(cell) for cell in row[:10] if cell])
    if any(keyword in str(row_str).lower() for keyword in ['plan', 'premium', 'current', 'total', 'monthly']):
        print(f'Row {ws._current_row-1}: {row_str}')

