import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import logging
import re
import argparse
import sys
from pathlib import Path
from validation import discrepancy_status, NOT_ON_CENSUS_STATUS

# Setup Logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Styling constants
# ---------------------------------------------------------------------------
_HEADER_FONT      = Font(name='Arial', bold=True, color='FFFFFF', size=10)
_HEADER_FILL      = PatternFill('solid', start_color='4472C4')
_CELL_FONT        = Font(name='Arial', size=10)
_CENTER           = Alignment(horizontal='center', vertical='center')
_LEFT             = Alignment(horizontal='left',   vertical='center')
_THIN             = Side(style='thin', color='D9D9D9')
_BORDER           = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_FILL_CORRECT     = PatternFill('solid', start_color='C6EFCE')   # green
_FILL_NOT_MATCH   = PatternFill('solid', start_color='FFC7CE')   # red
_FILL_NOT_INVOICE = PatternFill('solid', start_color='FFEB9C')   # yellow


class DynamicCensusFiller:
    def __init__(self):
        self.source_lookup = {}

        # Detected at runtime
        self.census_name_mode    = None   # 'split' | 'full'
        self.census_name_col     = None   # col index for full-name mode
        self.census_first_col    = None   # col index for split mode
        self.census_last_col     = None   # col index for split mode
        self.census_coverage_col = None
        self.plan_col            = None
        self.premium_col         = None
        self.discrepancy_col     = None
        self.header_row          = None
        self.data_start_row      = None

    # ------------------------------------------------------------------
    # Name normalisation
    # ------------------------------------------------------------------
    def normalize_name(self, name):
        """Intelligent name normalization for robust matching."""
        if not name or (isinstance(name, float) and pd.isna(name)): return ""
        
        s = str(name).lower().strip()
        
        # Handle "Last, First" format
        if ',' in s:
            parts = [p.strip() for p in s.split(',')]
            if len(parts) >= 2:
                s = f"{parts[1]} {parts[0]}"
                
        # Remove punctuation
        s = re.sub(r"[^a-z\s]", " ", s)
        s = re.sub(r"\s+", " ", s)
        
        parts = s.split()
        
        # Strip known suffixes
        strip_tokens = {'jr', 'sr', 'ii', 'iii', 'iv', 'v', 'esq', 'phd', 'md', 'dds', 'mr', 'mrs', 'ms', 'dr', 'prof'}
        
        cleaned_parts = []
        for p in parts:
            if p in strip_tokens:
                continue
            # Drop single-letter token (middle initial) only if name has 3+ tokens
            if len(p) == 1 and len(parts) >= 3:
                continue
            cleaned_parts.append(p)
            
        # Return sorted tokens so that FIRST LAST and LAST FIRST match exactly
        return " ".join(sorted(cleaned_parts))

    def is_valid_employee_name(self, raw_full_name):
        """Skip summary/total rows from source files."""
        text = str(raw_full_name or "").strip().lower()
        if not text:
            return False
        blocked_terms = ("total", "summary", "record", "employee details", "employer health plan")
        return not any(term in text for term in blocked_terms)

    # ------------------------------------------------------------------
    # Source loading
    # ------------------------------------------------------------------
    def find_source_columns(self, df):
        cols = {'first': None, 'last': None, 'full': None,
                'coverage': None, 'plan': None, 'premium': None}
        for col in df.columns:
            c = str(col).lower()
            if   'first'    in c and 'name' in c:                cols['first']    = col
            elif 'last'     in c and 'name' in c:                cols['last']     = col
            elif 'full'     in c and 'name' in c:                cols['full']     = col
            elif 'employee' in c and 'name' in c:                cols['full']     = col
            elif 'coverage' in c:                                cols['coverage'] = col
            elif 'plan'     in c and ('name' in c or 'desc' in c): cols['plan']   = col
            elif 'premium'  in c or 'current' in c:              cols['premium']  = col
        return cols

    def load_source(self, source_path):
        """Load invoice/source Excel and build a name-keyed lookup."""
        try:
            xl = pd.ExcelFile(source_path)
            sheet_name = next(
                (s for s in xl.sheet_names
                 if any(k in s.lower() for k in ('employee', 'detail', 'data'))),
                xl.sheet_names[0]
            )

            # Auto-detect header row (first row mentioning 'name' or 'plan')
            df_probe = pd.read_excel(source_path, sheet_name=sheet_name,
                                     nrows=10, header=None)
            header_row = 0
            for i, row in df_probe.iterrows():
                row_str = " ".join(str(x).lower() for x in row if pd.notna(x))
                if 'name' in row_str or 'plan' in row_str:
                    header_row = i
                    break

            df = pd.read_excel(source_path, sheet_name=sheet_name,
                               skiprows=header_row)
            logger.info(f"Loaded source sheet '{sheet_name}' with {len(df)} rows.")

            cols = self.find_source_columns(df)

            for _, row in df.iterrows():
                name_key = raw_full_name = ""

                if cols['full'] and pd.notna(row.get(cols['full'], None)):
                    raw_full_name = str(row[cols['full']]).strip()
                    name_key      = self.normalize_name(raw_full_name)
                elif cols['first'] and cols['last']:
                    f = str(row.get(cols['first'], '')).strip()
                    l = str(row.get(cols['last'],  '')).strip()
                    raw_full_name = f"{f} {l}".strip()
                    name_key      = self.normalize_name(raw_full_name)

                if not name_key:
                    continue
                if not self.is_valid_employee_name(raw_full_name):
                    continue

                premium = row[cols['premium']] if cols['premium'] else None
                if isinstance(premium, str):
                    premium = re.sub(r'[^\d.]', '', premium)
                    try:    premium = float(premium)
                    except: pass

                self.source_lookup[name_key] = {
                    'plan':     row[cols['plan']]     if cols['plan']     else None,
                    'premium':  premium,
                    'raw_name': raw_full_name,
                    'coverage': row[cols['coverage']] if cols['coverage'] else None,
                }

            logger.info(f"Source lookup built: {len(self.source_lookup)} entries.")
            return True

        except Exception as e:
            logger.error(f"Failed to load source: {e}")
            return False

    # ------------------------------------------------------------------
    # Census structure detection
    # ------------------------------------------------------------------
    def _write_header(self, ws, row, col, label):
        cell = ws.cell(row=row, column=col)
        cell.value     = label
        cell.font      = _HEADER_FONT
        cell.fill      = _HEADER_FILL
        cell.alignment = _CENTER
        cell.border    = _BORDER

    def _detect_census_structure(self, ws):
        """
        Scan the sheet for its header row and map all relevant columns.

        Supports two census name layouts:
          - 'full'  : a single 'Employee Name' (or 'Full Name') column
          - 'split' : separate 'First Name' and 'Last Name' columns

        If Plan Name / Monthly Premium / Discrepancies columns are absent
        they are appended automatically.
        """
        for r in range(1, 40):
            row_vals = {
                c: str(ws.cell(row=r, column=c).value or '').strip().lower()
                for c in range(1, ws.max_column + 1)
            }
            joined = " ".join(row_vals.values())
            if not any(k in joined for k in ('name', 'employee', 'first')):
                continue

            self.header_row     = r
            self.data_start_row = r + 1

            name_col = first_col = last_col = None
            coverage_col = plan_col = premium_col = disc_col = None

            for c, v in row_vals.items():
                if   'employee' in v and 'name' in v:   name_col     = c
                elif 'full'     in v and 'name' in v:   name_col     = c
                elif 'first'    in v and 'name' in v:   first_col    = c
                elif 'last'     in v and 'name' in v:   last_col     = c
                elif 'coverage' in v:                   coverage_col = c
                elif 'plan'     in v:                   plan_col     = c
                elif 'premium'  in v:                   premium_col  = c
                elif 'discrep'  in v:                   disc_col     = c

            # Determine name mode
            if name_col:
                self.census_name_mode = 'full'
                self.census_name_col  = name_col
            elif first_col and last_col:
                self.census_name_mode = 'split'
                self.census_first_col = first_col
                self.census_last_col  = last_col
            else:
                # Fallback: treat column 1 as a full-name column
                self.census_name_mode = 'full'
                self.census_name_col  = 1

            self.census_coverage_col = coverage_col

            # Append missing output columns
            max_col = ws.max_column

            if plan_col:
                self.plan_col = plan_col
            else:
                self.plan_col = max_col + 1
                self._write_header(ws, r, self.plan_col, 'Plan Name')
                max_col += 1

            if premium_col:
                self.premium_col = premium_col
            else:
                self.premium_col = max_col + 1
                self._write_header(ws, r, self.premium_col, 'Monthly Premium')
                max_col += 1

            if disc_col:
                self.discrepancy_col = disc_col
            else:
                self.discrepancy_col = max_col + 1
                self._write_header(ws, r, self.discrepancy_col, 'Discrepancies')

            # Column widths for output columns
            for col_idx, width in [
                (self.plan_col,         22),
                (self.premium_col,      18),
                (self.discrepancy_col,  15),
            ]:
                letter = ws.cell(row=1, column=col_idx).column_letter
                ws.column_dimensions[letter].width = width

            logger.info(
                f"Census structure — mode: {self.census_name_mode}, "
                f"header_row: {r}, coverage_col: {coverage_col}, "
                f"plan_col: {self.plan_col}, premium_col: {self.premium_col}, "
                f"disc_col: {self.discrepancy_col}"
            )
            return True

        logger.error("Could not detect header row in census template.")
        return False

    # ------------------------------------------------------------------
    # Template filling
    # ------------------------------------------------------------------
    def fill_template(self, template_path, output_path):
        """Fill plan, premium and discrepancy columns in the census template."""
        try:
            wb = load_workbook(template_path)

            # Accept any sheet — prefer one named census / table / employee
            ws = None
            for s in wb.sheetnames:
                if any(k in s.lower() for k in ('census', 'table', 'employee')):
                    ws = wb[s]
                    break
            if not ws:
                ws = wb.active

            if not self._detect_census_structure(ws):
                return False

            filled_count    = 0
            not_found_count = 0
            appended_count  = 0
            seen_census_names = set()
            last_data_row = self.data_start_row - 1

            for row_idx in range(self.data_start_row, ws.max_row + 1):

                # Extract employee name based on detected layout
                if self.census_name_mode == 'full':
                    raw_name = ws.cell(row=row_idx, column=self.census_name_col).value
                    if not raw_name:
                        break
                    emp_display = str(raw_name).strip()
                else:  # split
                    first = ws.cell(row=row_idx, column=self.census_first_col).value
                    last  = ws.cell(row=row_idx, column=self.census_last_col).value
                    if not first and not last:
                        if ws.cell(row=row_idx, column=1).value is None:
                            break
                        continue
                    emp_display = f"{first or ''} {last or ''}".strip()

                norm_name = self.normalize_name(emp_display)
                seen_census_names.add(norm_name)
                last_data_row = row_idx

                # Style output cells
                plan_cell = ws.cell(row=row_idx, column=self.plan_col)
                prem_cell = ws.cell(row=row_idx, column=self.premium_col)
                disc_cell = ws.cell(row=row_idx, column=self.discrepancy_col)

                for cell in (plan_cell, prem_cell, disc_cell):
                    cell.font   = _CELL_FONT
                    cell.border = _BORDER
                plan_cell.alignment = _LEFT
                prem_cell.alignment = _CENTER
                disc_cell.alignment = _CENTER

                # Match and fill
                if norm_name in self.source_lookup:
                    data     = self.source_lookup[norm_name]
                    coverage = (
                        ws.cell(row=row_idx, column=self.census_coverage_col).value
                        if self.census_coverage_col else None
                    )

                    status = discrepancy_status(
                        extracted_name=emp_display,
                        invoice_name=data['raw_name'],
                        extracted_coverage_tier=coverage,
                        invoice_coverage_tier=data['coverage'],
                        name_is_matched=True,
                    )

                    plan_cell.value = data['plan']

                    if data['premium'] is not None:
                        prem_cell.value         = data['premium']
                        prem_cell.number_format = '$#,##0.00'

                    disc_cell.value = status
                    disc_cell.fill  = _FILL_CORRECT if status == 'Correct' else _FILL_NOT_MATCH

                    filled_count += 1
                    logger.info(f"Matched & Filled: {emp_display} → {status}")

                else:
                    disc_cell.value = 'not available on invoice'
                    disc_cell.fill  = _FILL_NOT_INVOICE
                    not_found_count += 1
                    logger.debug(f"No match: {emp_display}")

            append_row = last_data_row
            for source_name_key, data in self.source_lookup.items():
                if source_name_key in seen_census_names:
                    continue

                append_row += 1

                # Name columns
                if self.census_name_mode == 'full':
                    ws.cell(row=append_row, column=self.census_name_col).value = data.get('raw_name')
                else:
                    raw_name = str(data.get('raw_name') or '').strip()
                    name_parts = raw_name.split()
                    first_name = name_parts[0] if name_parts else ''
                    last_name = " ".join(name_parts[1:]) if len(name_parts) > 1 else ''
                    ws.cell(row=append_row, column=self.census_first_col).value = first_name
                    ws.cell(row=append_row, column=self.census_last_col).value = last_name

                if self.census_coverage_col:
                    ws.cell(row=append_row, column=self.census_coverage_col).value = data.get('coverage')

                plan_cell = ws.cell(row=append_row, column=self.plan_col)
                prem_cell = ws.cell(row=append_row, column=self.premium_col)
                disc_cell = ws.cell(row=append_row, column=self.discrepancy_col)

                for cell in (plan_cell, prem_cell, disc_cell):
                    cell.font   = _CELL_FONT
                    cell.border = _BORDER
                plan_cell.alignment = _LEFT
                prem_cell.alignment = _CENTER
                disc_cell.alignment = _CENTER

                plan_cell.value = data.get('plan')
                if data.get('premium') is not None:
                    prem_cell.value = data.get('premium')
                    prem_cell.number_format = '$#,##0.00'

                disc_cell.value = NOT_ON_CENSUS_STATUS
                disc_cell.fill = _FILL_NOT_MATCH
                appended_count += 1

            wb.save(output_path)
            logger.info(
                f"Done! Saved to '{output_path}'. "
                f"Filled: {filled_count} | Not in Invoice: {not_found_count} | "
                f"Appended source-only: {appended_count}"
            )
            return True

        except Exception as e:
            logger.error(f"Failed to fill template: {e}")
            return False


def main():
    parser = argparse.ArgumentParser(description='Dynamic Insurance Census Filler')
    parser.add_argument('source',   help='Path to source/invoice Excel file')
    parser.add_argument('template', help='Path to census template Excel file')
    parser.add_argument('output', nargs='?', default='filled_output.xlsx',
                        help='Output filename (default: filled_output.xlsx)')
    args = parser.parse_args()

    filler = DynamicCensusFiller()
    if filler.load_source(args.source):
        if not filler.fill_template(args.template, args.output):
            sys.exit(1)
    else:
        sys.exit(1)

if __name__ == "__main__":
    main()