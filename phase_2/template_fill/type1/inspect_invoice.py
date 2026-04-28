import pandas as pd
import openpyxl

filename = 'invoice_pdf_20260305124606081-52796_8785628_729ab8_RPVE_20260408_124729_filtered.xlsx'

# List all sheets
wb = openpyxl.load_workbook(filename, read_only=True, data_only=True)
print('All sheets in data file (invoice):')
print(wb.sheetnames)
wb.close()

# Try reading first sheet
print('\n=== First sheet (index 0) ===')
try:
    df = pd.read_excel(filename, sheet_name=0)
    print('Columns:', list(df.columns))
    print('Shape:', df.shape)
    print('\\nHead:')
    print(df.head(3).to_string())
except Exception as e:
    print('Error:', e)
