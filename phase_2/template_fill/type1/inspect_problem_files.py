from openpyxl import load_workbook
import pandas as pd

def inspect_sheet(file_path, sheet_name, rows=25):
    print(f"\n--- {file_path} | Sheet: {sheet_name} ---")
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, nrows=rows)
        for i, row in df.iterrows():
            print(f"Row {i:2}: {row.tolist()[:15]}")
    except Exception as e:
        print(f"Error: {e}")

source = r'.\source_files\CDPHP_Invoice_-_East_Greenbush_RPVE_20260409_130013_filtered.xlsx'
template = r'.\census\2026 Engage Census  Network Validation East Greenbush Community Library MW.xlsx'

inspect_sheet(source, 'Employee Details')
inspect_sheet(template, 'Census', rows=40)
