import pandas as pd
import re
import logging

logger = logging.getLogger(__name__)

# --- Standardized Mapping Rules ---
CENSUS_COL_RULES = {
    'insured':   ['insured name', 'insured'],
    'emp_dep':   ['employee or dependent', 'emp or dep', 'relationship type', 'member type', 'relationship'],
    'coverage':  ['medical plan coverage', 'coverage level', 'coverage tier', 'enrollment status', 'medical:', 'medical', 'coverage'],
    'first':     ['first name', 'first', 'given name'],
    'last':      ['last name', 'last', 'surname'],
    'fullname':  ['full name', 'full_name', 'member name', 'employee name', 'name'],
    'gender':    ['gender', 'sex'],
    'dob':       ['date of birth', 'dob', 'birth date', 'birth', 'birthdate'],
    'zip':       ['zip code', 'zip', 'postal code', 'home zip'],
    'dep_rel':   ['dependent relation', 'dep relation'],
    'plan_desc': ['medical plan', 'plan description', 'plan name', 'plan'],
}

EMP_MARKERS = {
    'e', 'ee', 's', 'c', 'f', 'ec', 'es', 'ef', 'fam', 'nc', 'w', 'ne', 'ch', 
    'employee', 'subscriber', 'primary', 'insured', 'member', 'family',
    'employeefamily', 'employeeonly', 'employeechild', 'employeechildren', 'employeespouse',
    'employeeandspecial', 'employeeandchildren', 'employeeandspouse'
}
DEP_KEYWORDS = ('spouse', 'child', 'dependent', 'son', 'daughter', 'partner', 'domestic')

_HEADER_SIGNALS = (
    'first name', 'last name', 'date of birth', 'date of hire',
    'home zip', 'zip code', 'gender', 'dependent', 'enrollment',
    'employee/dependent', 'emp/dep', 'medical:', 'dental', 'vision',
    'coverage', 'relationship', 'subscriber',
)

def _is_valid_row(first, last, fullname):
    """Filters out summary or empty rows."""
    text = f"{first} {last} {fullname}".lower()
    if not text.strip(): return False
    blocked = ('total', 'summary', 'record', 'employee details', 'report')
    return not any(b in text for b in blocked)

def normalize_coverage_code(val):
    if not val or (isinstance(val, float) and pd.isna(val)): return ""
    t = re.sub(r'[\s+()\-]', '', str(val).upper())
    MAP = {
        'E': 'EE', 'S': 'ES', 'C': 'EC', 'F': 'FAM',
        'EE': 'EE', 'ES': 'ES', 'EC': 'EC', 'EF': 'FAM', 'FAM': 'FAM',
        'SP': 'ES', 'CH': 'EC', 'W': 'WO', 'NC': 'RC', 'NE': 'NE', 'C': 'C',
        'EMPLOYEE': 'EE', 'EMPLOYEEONLY': 'EE', 'EMPLOYEESPOUSE': 'ES',
        'EMPLOYEEANDSPOUSE': 'ES', 'EMPLOYEECHILDREN': 'EC', 'FAMILY': 'FAM'
    }
    return MAP.get(t, str(val).strip())

def detect_census_columns(df: pd.DataFrame) -> dict:
    col_map = {}
    cols = list(df.columns)
    for i, col in enumerate(cols):
        cl = str(col).lower().strip()
        for field, keywords in CENSUS_COL_RULES.items():
            if field in col_map: continue
            for kw in keywords:
                if kw in cl:
                    col_map[field] = col
                    # Special case: "Employee Name" might be followed by an unnamed first name column
                    if field == 'fullname' and i + 1 < len(cols):
                        next_col = cols[i+1]
                        if 'unnamed' in str(next_col).lower():
                            col_map['first_name_extra'] = next_col
                    break
    return col_map

def get_val(row, col_map, field, default=''):
    col = col_map.get(field)
    if not col: return default
    v = row.get(col, default)
    if v is None or (isinstance(v, float) and pd.isna(v)): return default
    return str(v).strip()

def normalize_census_to_list(path):
    """
    Standardizes any census Excel file into a list of Employee objects.
    Supports multi-sheet files (Employee sheet + Dependent sheet).
    """
    xl = pd.ExcelFile(path)
    
    # Identify Sheets
    emp_sheet = next((s for s in xl.sheet_names if any(k in s.lower() for k in ('census', 'employee', 'member', 'data'))), xl.sheet_names[0])
    dep_sheet = next((s for s in xl.sheet_names if s != emp_sheet and any(k in s.lower() for k in ('dependent', 'relative', 'family'))), None)
    
    # 1. Parse Employees
    df_emp = _load_best_sheet(path, emp_sheet)
    col_map_emp = detect_census_columns(df_emp)
    
    if 'insured' in col_map_emp and 'emp_dep' in col_map_emp:
        result = _parse_grouped(df_emp, col_map_emp)
    else:
        result = _parse_row_per_person(df_emp, col_map_emp)
    
    # 2. Parse Dependents from separate sheet if it exists
    if dep_sheet:
        _link_external_dependents(path, dep_sheet, result)
        
    return result

def _load_best_sheet(path, sheet_name):
    probe = pd.read_excel(path, sheet_name=sheet_name, header=None, nrows=30)
    best_hrow, best_score = 0, -1
    for i, row in probe.iterrows():
        row_str = ' '.join(str(v).lower() for v in row if pd.notna(v))
        score = sum(1 for sig in _HEADER_SIGNALS if sig in row_str)
        if score > best_score:
            best_score, best_hrow = score, i
    
    df = pd.read_excel(path, sheet_name=sheet_name, header=best_hrow)
    df.columns = [str(c).strip() for c in df.columns]
    return df

def _link_external_dependents(path, sheet_name, employees):
    """Parses a dedicated dependent sheet and links them to employees."""
    df = _load_best_sheet(path, sheet_name)
    col_map = detect_census_columns(df)
    
    # We need a way to link to employee. Usually 'Employee Name' or 'Subscriber'
    # Reuse 'insured' or 'fullname' keywords
    emp_link_col = col_map.get('insured') or col_map.get('fullname')
    dep_name_col = next((c for c in df.columns if any(k in str(c).lower() for k in ('dependant', 'dependent', 'child', 'spouse', 'member')) and c != emp_link_col), None)
    
    if not emp_link_col or not dep_name_col:
        logger.warning(f"Could not find link columns in dependent sheet '{sheet_name}'")
        return

    for _, row in df.iterrows():
        emp_raw = str(row.get(emp_link_col, '')).strip()
        dep_raw = str(row.get(dep_name_col, '')).strip()
        if not emp_raw or not dep_raw or 'nan' in emp_raw.lower() or 'nan' in dep_raw.lower():
            continue
            
        # Find employee
        from validation import normalize_text
        emp_key = normalize_text(emp_raw)
        
        # Try to find employee by various key formats
        target_emp = None
        for k, v in employees.items():
            # k is 'first last'
            if emp_key in k or k in emp_key:
                target_emp = v
                break
        
        if target_emp:
            # Parse dependent name
            d_parts = dep_raw.split(',')
            if len(d_parts) >= 2:
                d_last, d_first = d_parts[0].strip(), d_parts[1].strip()
            else:
                d_parts = dep_raw.split()
                d_first = d_parts[0] if d_parts else ''
                d_last = ' '.join(d_parts[1:]) if len(d_parts) > 1 else ''
            
            # DOB handling for dependents
            dob = None
            if 'dob' in col_map:
                dob = row.get(col_map['dob'])
            
            rel_val = str(row.get(col_map.get('dep_rel') or col_map.get('emp_dep') or '', '')).lower()
            rel = 'SP' if any(kw in rel_val for kw in ('spouse', 'wife', 'husband', 'partner')) else 'CH'
            
            target_emp['dependents'].append({
                'first': d_first, 'last': d_last, 
                'gender': _gender_code(get_val(row, col_map, 'gender')),
                'dob': dob if dob and str(dob).lower() != 'tbd' else None,
                'relation': rel
            })

def _parse_row_per_person(df, col_map):
    result = {}
    current_emp = None
    has_emp_dep = 'emp_dep' in col_map

    for _, row in df.iterrows():
        first = get_val(row, col_map, 'first')
        last = get_val(row, col_map, 'last')
        fullname = get_val(row, col_map, 'fullname')
        
        if not _is_valid_row(first, last, fullname): continue

        # Improved Name Splitting
        if not first and not last:
            extra = get_val(row, col_map, 'first_name_extra')
            if extra:
                last = fullname
                first = extra
            else:
                if ',' in fullname:
                    parts = [p.strip() for p in fullname.split(',')]
                    last = parts[0]
                    first = ' '.join(parts[1:])
                else:
                    parts = fullname.split()
                    if len(parts) >= 2:
                        first = parts[0]
                        last = ' '.join(parts[1:])
                    else:
                        first = parts[0] if parts else ''
                        last = ''
        
        if not first and not last: continue

        cov_raw = get_val(row, col_map, 'coverage')
        cov_clean = re.sub(r'[^a-z]', '', cov_raw.lower())
        
        # DOB Handling (Eland-style fix)
        dob, inferred_rel = None, None
        if 'dob' in col_map and pd.notna(row.get(col_map['dob'])):
            dob = row.get(col_map['dob'])
        else:
            for c_name, c_val in row.items():
                c_str = str(c_name).lower()
                if ('dob' in c_str or 'birth' in c_str) and pd.notna(c_val):
                    try:
                        ts = pd.Timestamp(c_val)
                        if not pd.isna(ts):
                            dob, inferred_rel = c_val, ('SP' if 'spouse' in c_str or 'partner' in c_str else 'CH')
                            break
                    except: pass

        # Employee/Dependent logic
        if has_emp_dep:
            emp_dep_val = get_val(row, col_map, 'emp_dep').lower().strip()
            is_employee = emp_dep_val in EMP_MARKERS
            is_dependent = emp_dep_val in DEP_KEYWORDS or emp_dep_val == '0'
            dep_relation = 'SP' if any(kw in emp_dep_val for kw in ('spouse', 'partner')) else 'CH'
        else:
            is_employee = cov_clean in EMP_MARKERS or (not cov_raw and current_emp is None)
            is_dependent = any(kw in cov_raw.lower() for kw in DEP_KEYWORDS) or (not cov_raw and current_emp is not None)
            dep_relation = inferred_rel

        if is_employee and not is_dependent:
            coverage = normalize_coverage_code(cov_raw) if cov_raw else 'EE'
            current_emp = {
                'first': first, 'last': last, 'gender': _gender_code(get_val(row, col_map, 'gender')),
                'dob': dob, 'zip': get_val(row, col_map, 'zip'), 'coverage': coverage, 'dependents': []
            }
            # Key uses first/last to avoid dups
            key = f"{first.lower()} {last.lower()}"
            if key not in result: result[key] = current_emp
        elif is_dependent and current_emp is not None:
            dep_last = last or current_emp['last']
            if not dep_relation: dep_relation = inferred_rel or 'CH'
            current_emp['dependents'].append({
                'first': first, 'last': dep_last, 'gender': _gender_code(get_val(row, col_map, 'gender')),
                'dob': dob, 'relation': dep_relation
            })
    return result

def _parse_grouped(df, col_map):
    # (Implementation similar to previous but using standardized helpers)
    from collections import defaultdict
    groups = defaultdict(list)
    for _, row in df.iterrows():
        insured = get_val(row, col_map, 'insured')
        if insured: groups[insured].append(row)
    
    result = {}
    for insured, rows in groups.items():
        emp_dep_col = col_map.get('emp_dep', '')
        emp_rows = [r for r in rows if str(r.get(emp_dep_col, '')).lower() in EMP_MARKERS]
        if not emp_rows: emp_rows = rows[:1]
        
        emp = emp_rows[0]
        first = get_val(emp, col_map, 'first')
        last = get_val(emp, col_map, 'last')
        fullname = get_val(emp, col_map, 'fullname')
        if not _is_valid_row(first, last, fullname): continue
        
        coverage = normalize_coverage_code(get_val(emp, col_map, 'coverage'))
        
        data = {
            'first': first, 'last': last, 'gender': _gender_code(get_val(emp, col_map, 'gender')),
            'dob': emp.get(col_map['dob']) if 'dob' in col_map else None,
            'zip': get_val(emp, col_map, 'zip'), 'coverage': coverage or 'EE',
            'dependents': []
        }
        
        dep_rows = [r for r in rows if r is not emp]
        for dr in dep_rows:
            d_first = get_val(dr, col_map, 'first')
            d_last = get_val(dr, col_map, 'last') or last
            rel_val = (get_val(dr, col_map, 'dep_rel') or get_val(dr, col_map, 'emp_dep')).lower()
            rel = 'SP' if any(kw in rel_val for kw in ('spouse', 'partner')) else 'CH'
            data['dependents'].append({
                'first': d_first, 'last': d_last, 'gender': _gender_code(get_val(dr, col_map, 'gender')),
                'dob': dr.get(col_map['dob']) if 'dob' in col_map else None,
                'relation': rel
            })
        
        key = f"{first.lower()} {last.lower()}"
        result[key] = data
    return result

def _gender_code(gender):
    if not gender: return ""
    g = str(gender).strip().upper()
    return 'M' if g.startswith('M') else ('F' if g.startswith('F') else '')
