"""
fill_template.py  –  3-input RAPT census filler
================================================
Usage:
    python fill_template.py <invoice> <ref_census> <rapt_template> [output]

Inputs
------
1. invoice        – BENEFITS_BILLING xlsx   (plan name, premium, coverage)
2. ref_census     – TEPCensus xlsx          (employee demographics + health plan coverage tier)
3. rapt_template  – RAPT_Census xlsx        (empty RAPT output template to fill)

Output
------
Filled RAPT xlsx with all RAPT columns including Discrepancies.
"""

import re
import logging
import argparse
from collections import defaultdict

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from validation import discrepancy_status, NOT_ON_CENSUS_STATUS

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Styling
# ---------------------------------------------------------------------------
_HEADER_FONT   = Font(name='Arial', bold=True, color='FFFFFF', size=10)
_HEADER_FILL   = PatternFill('solid', start_color='4472C4')
_CELL_FONT     = Font(name='Arial', size=10)
_CENTER        = Alignment(horizontal='center', vertical='center', wrap_text=False)
_LEFT          = Alignment(horizontal='left',   vertical='center')
_THIN          = Side(style='thin', color='D9D9D9')
_BORDER        = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_FILL_CORRECT  = PatternFill('solid', start_color='C6EFCE')
_FILL_MISMATCH = PatternFill('solid', start_color='FFC7CE')
_FILL_MISSING  = PatternFill('solid', start_color='FFEB9C')

# Name suffixes that should be ignored when building lookup keys
_SUFFIXES = {'jr', 'sr', 'ii', 'iii', 'iv', 'v', 'esq'}


# ---------------------------------------------------------------------------
# Name helpers
# ---------------------------------------------------------------------------
def _clean_name(name) -> str:
    """Lowercase, strip punctuation, collapse whitespace."""
    if not name or (isinstance(name, float) and pd.isna(name)):
        return ""
    s = str(name).lower().strip()
    if ',' in s:                          # "Last, First" → "first last"
        parts = [p.strip() for p in s.split(',')]
        s = f"{parts[1]} {parts[0]}" if len(parts) >= 2 else s
    s = re.sub(r'[^a-z0-9\s]', ' ', s)
    s = re.sub(r'\s+', ' ', s)
    return s.strip()


def _tokens(name) -> list:
    """Return significant name tokens, removing:
    - known suffixes (jr, sr, ii, iii, etc.) from ANY position
    - single-letter initials that are NOT the only token
    """
    raw = _clean_name(name)
    parts = raw.split()
    result = []
    for i, p in enumerate(parts):
        if p in _SUFFIXES:                          # drop suffix anywhere
            continue
        if len(p) == 1 and len(parts) > 2:          # drop lone initial if 3+ tokens
            continue
        result.append(p)
    return result


def _make_key(name) -> str:
    """
    Canonical lookup key: always stored as 'firstname lastname' order.
    Works for:
      - 'First Last'           → 'first last'
      - 'Last, First'          → 'first last'
      - 'LAST FIRST' (invoice) → stored under BOTH 'last first' AND 'first last'
        but _lookup_keys() returns both so either direction finds it.
    """
    return " ".join(_tokens(name))


def _lookup_keys(name) -> list:
    """
    Return all candidate lookup keys for a name.
    For a 2-token name both 'a b' and 'b a' are returned so
    'LAST FIRST' invoice names match 'First Last' census entries.
    For 3+ token names only the canonical form is returned (already handles
    middle names by stripping them in _tokens).
    """
    toks = _tokens(name)
    if not toks:
        return []
    key = " ".join(toks)
    keys = [key]
    if len(toks) == 2 and toks[0] != toks[1]:
        keys.append(f"{toks[1]} {toks[0]}")   # reversed
    return keys


def _is_valid_name(name) -> bool:
    text = str(name or '').strip().lower()
    if not text or text == 'nan':
        return False
    blocked = ('total', 'summary', 'record', 'employee details')
    return not any(b in text for b in blocked)


# ---------------------------------------------------------------------------
# Other helpers
# ---------------------------------------------------------------------------
def _plan_type_to_coverage(plan_type) -> str:
    if not plan_type or (isinstance(plan_type, float) and pd.isna(plan_type)):
        return ""
    t = re.sub(r'[\s+()\-]', '', str(plan_type).upper())
    MAP = {
        # Single-letter codes (CDPHP / Omni / RAPT census)
        'E':                 'EE',
        'S':                 'ES',
        'C':                 'EC',
        'F':                 'FAM',
        # Short codes
        'EE':                'EE',
        'ES':                'ES',
        'EC':                'EC',
        'EF':                'FAM',
        'FAM':               'FAM',
        'SP':                'ES',
        'CH':                'EC',
        # Long-form
        'EMPLOYEE':          'EE',
        'EMPLOYEEONLY':      'EE',
        'EMPLOYEESPOUSE':    'ES',
        'EMPLOYEEANDSPOUSE': 'ES',
        'EMPLOYEECHILDREN':  'EC',
        'EMPLOYEECHILDRN':   'EC',
        'EMPLOYEEANDCHILDREN': 'EC',
        'FAMILY':            'FAM',
        'EMPLOYEEFAMILY':    'FAM',
    }
    # Return mapped value; if not in map, return the original value as-is
    # (do NOT default to EE — that hides real data)
    return MAP.get(t, str(plan_type).strip())


def _gender_code(gender) -> str:
    if not gender or (isinstance(gender, float) and pd.isna(gender)):
        return ""
    g = str(gender).strip().upper()
    return 'M' if g.startswith('M') else ('F' if g.startswith('F') else '')


def _relationship_code(emp_or_dep: str, relation) -> str:
    if str(emp_or_dep).lower() == 'employee':
        return 'EE'
    return 'SP' if 'spouse' in str(relation or '').lower() else 'CH'


# ---------------------------------------------------------------------------
# Step 1: Load invoice (BENEFITS_BILLING)
# ---------------------------------------------------------------------------
def load_invoice(path: str) -> dict:
    """
    Returns dict keyed by ALL candidate lookup keys for each employee name.
    BENEFITS_BILLING positional layout (after header row):
      col0=full_name, col1=first, col2=middle, col3=last,
      col4=coverage,  col5=plan_name, col6=plan_type,
      col7=current_premium, col8=adj, col9=birth, col10=gender,
      col11=home_zip, col12=billing_period
    """
    xl = pd.ExcelFile(path)
    sheet = next(
        (s for s in xl.sheet_names if any(k in s.lower() for k in ('employee', 'detail', 'data'))),
        xl.sheet_names[0]
    )

    # Detect header row
    probe = pd.read_excel(path, sheet_name=sheet, nrows=10, header=None)
    hrow = 0
    for i, row in probe.iterrows():
        if any(str(v).lower() in ('full name', 'first name', 'plan name') for v in row if pd.notna(v)):
            hrow = i
            break

    df = pd.read_excel(path, sheet_name=sheet, skiprows=hrow + 1, header=None)
    lookup = {}

    for _, row in df.iterrows():
        row_list = list(row)
        raw_full = str(row_list[0] if row_list else '').strip()
        if not raw_full or not _is_valid_name(raw_full):
            continue

        coverage_raw = row_list[4]  if len(row_list) > 4  else None
        plan_raw     = row_list[5]  if len(row_list) > 5  else None
        premium_raw  = row_list[7]  if len(row_list) > 7  else None
        zip_raw      = row_list[11] if len(row_list) > 11 else None

        if isinstance(premium_raw, str):
            premium_raw = re.sub(r'[^\d.]', '', premium_raw)
            try:    premium_raw = float(premium_raw)
            except: premium_raw = None
        elif isinstance(premium_raw, float) and pd.isna(premium_raw):
            premium_raw = None

        zip_val = str(zip_raw).strip() if zip_raw and pd.notna(zip_raw) else ''

        entry = {
            'plan':     plan_raw,
            'premium':  premium_raw,
            'raw_name': raw_full,
            'coverage': coverage_raw,
            'zip':      zip_val,
        }

        # Store under every candidate key so LAST FIRST and First Last both resolve
        for key in _lookup_keys(raw_full):
            if key:
                lookup[key] = entry

    unique = len({v['raw_name'] for v in lookup.values()})
    logger.info(f"Invoice loaded: {unique} employees from '{sheet}'")
    return lookup


# ---------------------------------------------------------------------------
# Step 2: Universal reference census loader (any format)
# ---------------------------------------------------------------------------

# ── Column keyword rules ────────────────────────────────────────────────────
# For each logical field, any column whose header contains ANY of these
# keywords (case-insensitive) will be mapped to that field.
_CENSUS_COL_RULES = {
    # ── Ordered from MOST specific to LEAST specific ──────────────────────────
    # Each entry is (logical_field, list_of_keywords_in_priority_order).
    # The FIRST column whose header CONTAINS any keyword wins.

    'insured':   ['insured name', 'insured'],                   # grouping key (TEPCensus)

    'emp_dep':   ['employee or dependent', 'emp or dep',        # explicit employee/dep marker
                  'employee (1)', 'dependent (0)',
                  'relationship type', 'member type',
                  'relationship'],                               # CDPHP: 'Relationship' col

    # Coverage/tier — MOST SPECIFIC keywords first so 'Medical Plan Coverage'
    # wins before the generic 'coverage' substring catches 'Relationship'.
    'coverage':  ['medical plan coverage',                      # CDPHP Employee Census
                  'coverage level', 'coverage tier',
                  'coverage type', 'benefit tier',
                  'plan coverage',                              # generic "plan coverage"
                  'coverage'],                                  # last resort — generic
    # NOTE: 'relationship' and 'plan type' deliberately removed from coverage
    # to avoid ambiguity with the Relationship and Medical Plan columns.

    'first':     ['first name', 'first', 'given name'],        # first name ONLY
    'last':      ['last name', 'last', 'surname', 'family name'],
    'fullname':  ['full name', 'full_name', 'member name',
                  'employee name', 'name'],                    # combined first+last

    'gender':    ['gender', 'sex'],
    'dob':       ['date of birth', 'dob', 'birth date',
                  'birth', 'birthdate'],
    'zip':       ['zip code', 'zip', 'postal code', 'home zip'],
    'dep_rel':   ['dependent relation', 'dep relation'],        # 'relation' removed; covered by emp_dep
    'plan_desc': ['medical plan', 'plan description', 'plan name', 'plan'],
}

# Employee-type values — used when coverage column holds LONG-FORM text like
# 'Employee Only' / 'Employee + Spouse', OR when the Relationship column
# (emp_dep) holds 'Employee'.
_EMP_MARKERS = {
    # Single-letter codes (Omni / RAPT census)
    'e', 'ee', 's', 'c', 'f', 'ec', 'es', 'ef', 'fam',
    # Long-form relationship values (CDPHP Employee Census 'Relationship' column)
    'employee',
    # Generic
    'subscriber', 'primary', 'insured', 'member',
}
# Dependent relation keywords
_DEP_KEYWORDS = ('spouse', 'child', 'dependent', 'son', 'daughter',
                 'partner', 'domestic')



def _detect_census_columns(df: pd.DataFrame) -> dict:
    """
    Returns a mapping {logical_field: actual_column_name} by scanning
    column headers with keyword rules. First match wins per field.
    """
    col_map = {}
    for col in df.columns:
        cl = col.lower().strip()
        for field, keywords in _CENSUS_COL_RULES.items():
            if field in col_map:
                continue  # already mapped
            for kw in keywords:
                if kw in cl:
                    col_map[field] = col
                    break
    return col_map


def _get_val(row, col_map: dict, field: str, default='') -> str:
    """Safely get a string value from a row using the col_map."""
    col = col_map.get(field)
    if not col:
        return default
    v = row.get(col, default)
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return default
    return str(v).strip()


def load_ref_census(path: str) -> dict:
    """
    Universal census loader — works with ANY column naming convention.

    Strategy:
      1. Find the right sheet
      2. Find the header row (first row that mentions a name-related keyword)
      3. Map columns by keyword detection
      4. Detect layout type:
         A) GROUPED  – has an 'insured' column that groups members together
                       (e.g. TEPCensus: Insured Name + Employee or Dependent)
         B) ROW-PER-PERSON – each row = one person, employee vs dependent
                       identified by coverage code or relation keyword
                       (e.g. Omni: Coverage Level = E/F/C or 'Spouse'/'Child')
         C) EMPLOYEE-ONLY – no dependency structure, every row = one employee
    Returns dict keyed by canonical 'firstname lastname' lookup key.
    """
    xl = pd.ExcelFile(path)

    # Pick the best sheet
    sheet = next(
        (s for s in xl.sheet_names
         if any(k in s.lower() for k in ('census', 'employee', 'member', 'data',
                                          'enrollment', 'roster'))),
        xl.sheet_names[0]
    )

    # Find header row (first row containing a name / employee keyword)
    probe = pd.read_excel(path, sheet_name=sheet, header=None, nrows=20)
    hrow = 0
    for i, row in probe.iterrows():
        row_str = ' '.join(str(v).lower() for v in row if pd.notna(v))
        if any(kw in row_str for kw in ('first', 'last', 'insured', 'name',
                                         'employee', 'member', 'subscriber')):
            hrow = i
            break

    df = pd.read_excel(path, sheet_name=sheet, header=hrow)
    df.columns = [str(c).strip() for c in df.columns]

    col_map = _detect_census_columns(df)
    logger.info(f"Census column map: {col_map}")

    # ── Layout detection ────────────────────────────────────────────────────
    if 'insured' in col_map and 'emp_dep' in col_map:
        result = _parse_grouped(df, col_map)
        fmt = 'grouped (TEPCensus-style)'
    elif 'first' in col_map or 'fullname' in col_map:
        result = _parse_row_per_person(df, col_map)
        fmt = 'row-per-person (Omni-style)'
    else:
        logger.warning("Census: no recognisable name column found — returning empty.")
        return {}

    logger.info(f"Ref census loaded ({fmt}): {len(result)} unique employees")
    return result


def _parse_grouped(df: pd.DataFrame, col_map: dict) -> dict:
    """
    Grouped layout: rows share an 'Insured Name' grouping key.
    Employee row identified by emp_dep column == 'employee'.
    """
    groups = defaultdict(list)
    for _, row in df.iterrows():
        insured = _get_val(row, col_map, 'insured')
        if insured:
            groups[insured].append(row)

    result = {}
    for insured, rows in groups.items():
        emp_dep_col = col_map.get('emp_dep', '')
        emp_rows = [r for r in rows
                    if str(r.get(emp_dep_col, '') or '').lower() == 'employee']
        if not emp_rows:
            emp_rows = rows[:1]  # fallback: treat first row as employee

        emp      = emp_rows[0]
        first    = _get_val(emp, col_map, 'first')
        last     = _get_val(emp, col_map, 'last')
        gender   = _gender_code(_get_val(emp, col_map, 'gender'))
        dob      = emp.get(col_map['dob']) if 'dob' in col_map else None
        zip_code = _get_val(emp, col_map, 'zip')

        # Coverage from plan rows
        plan_desc_col = col_map.get('plan_desc', '')
        health_rows   = [r for r in emp_rows
                         if 'BCBS' in str(r.get(plan_desc_col, '') or '').upper()]
        if health_rows:
            coverage = _plan_type_to_coverage(
                _get_val(health_rows[0], col_map, 'coverage'))
        else:
            raw_cov  = _get_val(emp, col_map, 'coverage')
            coverage = _plan_type_to_coverage(raw_cov) if raw_cov else 'EE'

        # Dependents
        dep_rows   = [r for r in rows if r is not emp and
                      str(r.get(emp_dep_col, '') or '').lower() != 'employee']
        dependents = _extract_dependents(dep_rows, col_map, zip_code)

        canon_key = _make_key(f"{first} {last}")
        if canon_key and canon_key not in result:
            result[canon_key] = {
                'first': first, 'last': last, 'gender': gender,
                'dob': dob, 'zip': zip_code, 'coverage': coverage,
                'dependents': dependents,
            }
    return result


def _parse_row_per_person(df: pd.DataFrame, col_map: dict) -> dict:
    """
    Row-per-person layout: each row is one person.
    Employees identified by:
      1. emp_dep column (if available) — e.g. 'Employee' / 'Spouse' / 'Child'
      2. Fallback: coverage code or absence of relation keyword
    Dependents follow the preceding employee row.
    """
    result      = {}
    current_emp = None
    has_emp_dep = 'emp_dep' in col_map   # explicit Employee/Dependent marker column

    for _, row in df.iterrows():
        # Get name — prefer first+last split, fall back to fullname
        first = _get_val(row, col_map, 'first')
        last  = _get_val(row, col_map, 'last')

        if not first and not last:
            fullname = _get_val(row, col_map, 'fullname')
            if not fullname or fullname.lower() == 'nan':
                continue
            parts = fullname.split()
            first = parts[0] if parts else ''
            last  = ' '.join(parts[1:]) if len(parts) > 1 else ''

        if not first and not last:
            continue

        cov_raw   = _get_val(row, col_map, 'coverage')
        cov_lower = cov_raw.lower()
        cov_clean = re.sub(r'[^a-z]', '', cov_lower)

        gender   = _gender_code(_get_val(row, col_map, 'gender'))
        dob      = row.get(col_map['dob']) if 'dob' in col_map else None
        zip_code = _get_val(row, col_map, 'zip')

        # ── Employee / Dependent classification ──────────────────────
        if has_emp_dep:
            # PRIMARY: use explicit marker column (most reliable)
            emp_dep_val = _get_val(row, col_map, 'emp_dep').lower().strip()
            is_employee  = emp_dep_val in ('employee', '1', 'ee', 'subscriber',
                                           'primary', 'insured', 'member')
            is_dependent = emp_dep_val in ('spouse', 'child', 'dependent', '0',
                                           'sp', 'ch', 'son', 'daughter',
                                           'partner', 'domestic partner')
            # Derive relation from emp_dep value directly
            dep_relation = ('SP' if any(kw in emp_dep_val
                            for kw in ('spouse', 'partner')) else 'CH')
        else:
            # FALLBACK: infer from coverage codes
            is_employee  = (
                cov_clean in _EMP_MARKERS
                or (not cov_raw and current_emp is None)
            )
            is_dependent = (
                any(kw in cov_lower for kw in _DEP_KEYWORDS)
                or (_get_val(row, col_map, 'dep_rel') != '' and not is_employee)
                or (not cov_raw and current_emp is not None)
            )
            dep_relation = None   # will be derived below

        is_wc_only = cov_clean in ('wc', 'workcomp', 'workerscomp')

        if is_wc_only:
            current_emp = None
            continue

        if is_employee and not is_dependent:
            coverage = _plan_type_to_coverage(cov_raw) if cov_raw else 'EE'
            current_emp = {
                'first': first, 'last': last, 'gender': gender,
                'dob': dob, 'zip': zip_code, 'coverage': coverage,
                'dependents': [],
            }
            canon_key = _make_key(f"{first} {last}")
            if canon_key and canon_key not in result:
                result[canon_key] = current_emp

        elif is_dependent and current_emp is not None:
            dep_last = last or current_emp['last']
            if dep_relation is None:
                # Fallback relation derivation from dep_rel or coverage
                rel_val  = _get_val(row, col_map, 'dep_rel') or cov_raw
                dep_relation = ('SP' if any(kw in rel_val.lower()
                                for kw in ('spouse', 'partner')) else 'CH')
            if not any(d['first'] == first and d['last'] == dep_last
                       for d in current_emp['dependents']):
                current_emp['dependents'].append({
                    'first': first, 'last': dep_last,
                    'gender': gender, 'dob': dob,
                    'zip': current_emp['zip'], 'relation': dep_relation,
                })

    return result


def _extract_dependents(dep_rows, col_map: dict, emp_zip: str) -> list:
    """Build dependent list from a set of rows (grouped-layout helper)."""
    dependents, seen = [], set()
    for dr in dep_rows:
        d_first  = _get_val(dr, col_map, 'first')
        d_last   = _get_val(dr, col_map, 'last')
        dep_key  = f"{d_first}|{d_last}"
        if dep_key in seen or (not d_first and not d_last):
            continue
        seen.add(dep_key)
        rel_val  = _get_val(dr, col_map, 'dep_rel') or _get_val(dr, col_map, 'coverage')
        rel      = 'SP' if any(kw in rel_val.lower()
                                for kw in ('spouse', 'partner')) else 'CH'
        dependents.append({
            'first':    d_first, 'last':    d_last,
            'gender':   _gender_code(_get_val(dr, col_map, 'gender')),
            'dob':      dr.get(col_map['dob']) if 'dob' in col_map else None,
            'zip':      emp_zip,
            'relation': rel,
        })
    return dependents


# ---------------------------------------------------------------------------
# Step 3: Detect RAPT template column layout
# ---------------------------------------------------------------------------
def _col_key(val) -> str:
    return re.sub(r'\s+', ' ', str(val or '').replace('*', '').strip().lower())


def detect_rapt_columns(ws) -> dict:
    for r in range(1, 25):
        row_keys = {c: _col_key(ws.cell(row=r, column=c).value)
                    for c in range(1, ws.max_column + 1)}
        if 'first name' not in row_keys.values():
            continue

        cols = {'header_row': r, 'data_start': r + 1}
        for c, key in row_keys.items():
            if   key == 'data row':                          cols['data_row'] = c
            elif key == 'first name':                        cols['first']    = c
            elif key == 'last name':                         cols['last']     = c
            elif 'gender' in key:                            cols['gender']   = c
            elif 'date of birth' in key or key == 'birth':  cols['dob']      = c
            elif 'zip' in key:                               cols['zip']      = c
            elif 'relationship' in key:                      cols['relation'] = c
            elif 'dependent of' in key:                      cols['dep_of']   = c
            elif 'coverage' in key and 'cobra' not in key:  cols['coverage'] = c
            elif 'cobra' in key:                             cols['cobra']    = c
            elif key == 'plan name':                         cols['plan']     = c
            elif 'monthly' in key and 'premium' in key:     cols['premium']  = c
            elif 'discrepanc' in key:                        cols['disc']     = c

        # Append Discrepancies column if absent
        if 'disc' not in cols:
            dc   = (cols.get('premium') or ws.max_column) + 1
            cell = ws.cell(row=r, column=dc)
            cell.value     = 'Discrepancies'
            cell.font      = _HEADER_FONT
            cell.fill      = _HEADER_FILL
            cell.alignment = _CENTER
            cell.border    = _BORDER
            ws.column_dimensions[cell.column_letter].width = 22
            cols['disc'] = dc

        # Column widths
        for k, w in [('plan', 32), ('premium', 18), ('first', 15), ('last', 18), ('dob', 14)]:
            if k in cols:
                ws.column_dimensions[ws.cell(row=1, column=cols[k]).column_letter].width = w

        logger.info(f"RAPT columns: {cols}")
        return cols

    raise ValueError("RAPT header row not found (no 'First Name' column in first 25 rows)")


# ---------------------------------------------------------------------------
# Step 4: Write rows into RAPT template
# ---------------------------------------------------------------------------
def _wcell(ws, row, col, value, align=None, fmt=None, fill=None):
    if col is None:
        return
    cell         = ws.cell(row=row, column=col)
    cell.value   = value
    cell.font    = _CELL_FONT
    cell.border  = _BORDER
    if align: cell.alignment     = align
    if fmt:   cell.number_format = fmt
    if fill:  cell.fill          = fill


def _write_dob(ws, row, col, dob):
    if col is None or dob is None:
        return
    try:
        ts = pd.Timestamp(dob)
        if pd.isna(ts):
            return
        cell               = ws.cell(row=row, column=col)
        cell.value         = ts.to_pydatetime()
        cell.font          = _CELL_FONT
        cell.border        = _BORDER
        cell.alignment     = _CENTER
        cell.number_format = 'MM/DD/YYYY'
    except Exception:
        pass


def write_employee_row(ws, row_idx, data_row_num, emp, inv, cols):
    zip_val = (inv.get('zip') or emp.get('zip') or '') if inv else emp.get('zip', '')

    _wcell(ws, row_idx, cols.get('data_row'), data_row_num,    _CENTER)
    _wcell(ws, row_idx, cols.get('first'),    emp['first'],    _LEFT)
    _wcell(ws, row_idx, cols.get('last'),     emp['last'],     _LEFT)
    _wcell(ws, row_idx, cols.get('gender'),   emp['gender'],   _CENTER)
    _write_dob(ws, row_idx, cols.get('dob'),  emp.get('dob'))
    _wcell(ws, row_idx, cols.get('zip'),      zip_val,         _CENTER)
    _wcell(ws, row_idx, cols.get('relation'), 'EE',            _CENTER)
    _wcell(ws, row_idx, cols.get('dep_of'),   '',              _CENTER)
    _wcell(ws, row_idx, cols.get('coverage'), emp['coverage'], _CENTER)
    _wcell(ws, row_idx, cols.get('cobra'),    'N',             _CENTER)

    if inv:
        _wcell(ws, row_idx, cols.get('plan'), inv['plan'], _LEFT)
        if inv.get('premium') is not None:
            _wcell(ws, row_idx, cols.get('premium'), inv['premium'], _CENTER, '$#,##0.00')

        status = discrepancy_status(
            extracted_name          = f"{emp['first']} {emp['last']}",
            invoice_name            = inv['raw_name'],
            extracted_coverage_tier = emp['coverage'],
            invoice_coverage_tier   = inv['coverage'],
        )
        _wcell(ws, row_idx, cols.get('disc'), status, _CENTER,
               fill=_FILL_CORRECT if status == 'Correct' else _FILL_MISMATCH)
    else:
        _wcell(ws, row_idx, cols.get('disc'), 'not available on invoice',
               _CENTER, fill=_FILL_MISSING)


def write_dependent_row(ws, row_idx, data_row_num, dep, emp_row_num, cols):
    zip_val = dep.get('zip', '')
    _wcell(ws, row_idx, cols.get('data_row'), data_row_num,    _CENTER)
    _wcell(ws, row_idx, cols.get('first'),    dep['first'],    _LEFT)
    _wcell(ws, row_idx, cols.get('last'),     dep['last'],     _LEFT)
    _wcell(ws, row_idx, cols.get('gender'),   dep['gender'],   _CENTER)
    _write_dob(ws, row_idx, cols.get('dob'),  dep.get('dob'))
    _wcell(ws, row_idx, cols.get('zip'),      zip_val,         _CENTER)
    _wcell(ws, row_idx, cols.get('relation'), dep['relation'], _CENTER)
    _wcell(ws, row_idx, cols.get('dep_of'),   emp_row_num,     _CENTER)
    # Coverage / Cobra / Plan / Premium / Disc intentionally blank for dependents


# ---------------------------------------------------------------------------
# Main orchestrator
# ---------------------------------------------------------------------------
def fill_rapt_template(invoice_path, ref_census_path, template_path, output_path):
    invoice_lookup = load_invoice(invoice_path)
    ref_lookup     = load_ref_census(ref_census_path)

    wb = load_workbook(template_path)
    ws = next(
        (wb[s] for s in wb.sheetnames
         if any(k in s.lower() for k in ('sheet', 'census', 'table', 'employee'))),
        wb.active
    )

    cols         = detect_rapt_columns(ws)
    write_row    = cols['data_start']
    data_row_num = 1

    matched = not_on_invoice = not_on_ref = 0
    seen_invoice_keys = set()   # canonical keys already written

    # --- Rows from reference census ---
    for canon_key, emp in ref_lookup.items():
        # Try all lookup-key variants to find invoice match
        inv = None
        for k in _lookup_keys(f"{emp['first']} {emp['last']}"):
            if k in invoice_lookup:
                inv = invoice_lookup[k]
                seen_invoice_keys.add(k)
                # Also mark the reversed key as seen to prevent duplicate append
                toks = k.split()
                if len(toks) == 2:
                    seen_invoice_keys.add(f"{toks[1]} {toks[0]}")
                break

        if inv:
            matched += 1
        else:
            not_on_invoice += 1

        emp_row_num = data_row_num
        write_employee_row(ws, write_row, data_row_num, emp, inv, cols)
        write_row += 1; data_row_num += 1

        for dep in emp.get('dependents', []):
            write_dependent_row(ws, write_row, data_row_num, dep, emp_row_num, cols)
            write_row += 1; data_row_num += 1

    # --- Invoice-only rows (on invoice but NOT in ref census) ---
    for inv_key, inv in invoice_lookup.items():
        if inv_key in seen_invoice_keys:
            continue
        # Skip if the reversed key was already written
        toks = inv_key.split()
        if len(toks) == 2 and f"{toks[1]} {toks[0]}" in seen_invoice_keys:
            continue

        not_on_ref += 1
        seen_invoice_keys.add(inv_key)   # prevent processing same person twice

        raw    = str(inv['raw_name']).strip()
        c_name = _clean_name(raw)
        t      = [x for x in c_name.split() if x not in _SUFFIXES]
        # Invoice is LAST FIRST — swap to First Last for the RAPT output
        if len(t) >= 2:
            first, last = t[-1].title(), t[0].title()
        else:
            first, last = (t[0].title() if t else raw), ''

        fake_emp = {
            'first': first, 'last': last, 'gender': '',
            'dob': None, 'zip': inv.get('zip', ''),
            'coverage': inv.get('coverage', ''),
            'dependents': [],
        }
        write_employee_row(ws, write_row, data_row_num, fake_emp, inv, cols)

        # Override discrepancy to NOT_ON_CENSUS_STATUS
        dc = ws.cell(row=write_row, column=cols['disc'])
        dc.value = NOT_ON_CENSUS_STATUS
        dc.fill  = _FILL_MISMATCH
        dc.alignment = _CENTER
        dc.font  = _CELL_FONT
        dc.border = _BORDER

        write_row += 1; data_row_num += 1

    wb.save(output_path)
    logger.info(
        f"Saved '{output_path}'.  "
        f"Matched={matched} | Not on invoice={not_on_invoice} | "
        f"Invoice-only (appended)={not_on_ref}"
    )


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(
        description='Fill RAPT census template from invoice + reference census'
    )
    parser.add_argument('invoice',    help='BENEFITS_BILLING xlsx')
    parser.add_argument('ref_census', help='TEPCensus xlsx')
    parser.add_argument('template',   help='RAPT_Census xlsx (empty template)')
    parser.add_argument('output', nargs='?', default='filled_rapt_output.xlsx')
    args = parser.parse_args()
    fill_rapt_template(args.invoice, args.ref_census, args.template, args.output)


if __name__ == '__main__':
    main()