"""
data_validation.py  —  Unified Post-Fill Name Normalisation & Discrepancy Resolver
====================================================================================
Last step in the pipeline — runs AFTER fill_template.py (type1 / type2 / type3).

Pipeline:
  Phase 1  →  Invoice PDF → extraction.xlsx   (raw invoice data)
  Phase 2  →  fill_template.py → filled_census.xlsx   (filled with strict name match)
  Phase 3  →  data_validation.py → validated_census.xlsx  (THIS FILE — fuzzy resolve)

What it does:
  1. Reads the filled Excel (output of fill_template).
  2. Also reads the original Phase 1 invoice Excel to get all invoice names.
  3. For every row still flagged as "Not on census" or "not available on invoice":
       a. Normalises both names (strips middle init, handles LAST FIRST / First Last)
       b. Pass 1 — exact canonical match
       c. Pass 2 — token-swap match (LAST FIRST  ↔  First Last)
       d. Pass 3 — fuzzy similarity match (threshold configurable)
  4. If a match is found: fills in plan/premium and updates Discrepancy cell.
  5. Saves a new validated Excel + JSON audit log.

Usage:
    python data_validation.py <filled_excel> <invoice_excel> [output] [--threshold 85]

    filled_excel   — output Excel from fill_template.py
    invoice_excel  — Phase 1 extraction Excel (raw invoice data)
    output         — (optional) output path, default: <filled>_validated.xlsx
    --threshold    — fuzzy match minimum score 0-100 (default 85)
"""

from __future__ import annotations

import argparse
import json
import logging
import re
import sys
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

# ---------------------------------------------------------------------------
# Optional fuzzy library — falls back to SequenceMatcher if not installed
# ---------------------------------------------------------------------------
try:
    from rapidfuzz import fuzz as _fuzz
    def _similarity(a: str, b: str) -> float:
        return _fuzz.token_sort_ratio(a, b)
except ImportError:
    try:
        from fuzzywuzzy import fuzz as _fuzz  # type: ignore
        def _similarity(a: str, b: str) -> float:
            return _fuzz.token_sort_ratio(a, b)
    except ImportError:
        from difflib import SequenceMatcher
        def _similarity(a: str, b: str) -> float:  # type: ignore
            return SequenceMatcher(None, a, b).ratio() * 100

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Styling
# ---------------------------------------------------------------------------
_FONT         = Font(name='Arial', size=10)
_FONT_BOLD    = Font(name='Arial', size=10, bold=True)
_CENTER       = Alignment(horizontal='center', vertical='center')
_LEFT         = Alignment(horizontal='left',   vertical='center')
_FILL_GREEN   = PatternFill('solid', start_color='C6EFCE')   # correct
_FILL_YELLOW  = PatternFill('solid', start_color='FFEB9C')   # fuzzy / uncertain
_FILL_RED     = PatternFill('solid', start_color='FFC7CE')   # still unresolved
_FILL_ORANGE  = PatternFill('solid', start_color='FFD966')   # possible match

# Discrepancy status strings (matching what fill_template uses)
_NOT_ON_CENSUS   = "Not on census"
_NOT_ON_INVOICE  = "not available on invoice"
_CORRECT         = "Correct"

# Known suffixes / titles to strip from names
_STRIP_TOKENS = {
    'jr', 'sr', 'ii', 'iii', 'iv', 'v', 'esq', 'phd', 'md', 'dds',
    'mr', 'mrs', 'ms', 'dr', 'prof',
}


# ===========================================================================
# NAME NORMALISATION ENGINE
# ===========================================================================

def _clean(raw) -> str:
    """Lowercase, remove punctuation except spaces, collapse whitespace."""
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return ""
    s = str(raw).strip().lower()
    # Handle "Last, First" comma format → "first last"
    if ',' in s:
        parts = [p.strip() for p in s.split(',')]
        if len(parts) >= 2:
            s = f"{parts[1]} {parts[0]}"
    s = re.sub(r"[^a-z\s]", " ", s)   # keep only letters and spaces
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def _tokens(raw) -> list[str]:
    """
    Return significant name tokens after:
      - Lowercasing and punctuation removal
      - Dropping known suffixes/titles
      - Dropping single-letter middle initials (when ≥ 3 tokens present)
    """
    cleaned = _clean(raw)
    parts = cleaned.split()
    result = []
    for p in parts:
        if p in _STRIP_TOKENS:
            continue
        if len(p) == 1 and len(parts) >= 3:
            # Drop single-letter token (middle initial) only if name has 3+ tokens
            continue
        result.append(p)
    return result


def canonical(raw) -> str:
    """
    Return canonical form: always 'firstname lastname' order, stripped of
    middle initials and suffixes.

    Examples:
        'BAKER HOLLI M'   → 'holli baker'  (token-swap because invoice is LAST FIRST)
        'Holli Baker'     → 'holli baker'
        'Baker, Holli M'  → 'holli baker'
        'LORENZ LUCAS'    → 'lucas lorenz'  (swap)
    """
    toks = _tokens(raw)
    if not toks:
        return ""
    # We store canonical as 'first last' — but since we don't know
    # which is first, we return ALL permutations via lookup_keys()
    return " ".join(toks)


def lookup_keys(raw) -> list[str]:
    """
    Return all candidate lookup keys for a name:
      - Original token order
      - Reversed (for LAST FIRST ↔ First Last resolution)
    For 3+ token names after stripping, all permutations of first/last pair.
    """
    toks = _tokens(raw)
    if not toks:
        return []
    key = " ".join(toks)
    keys = [key]
    if len(toks) == 2 and toks[0] != toks[1]:
        keys.append(f"{toks[1]} {toks[0]}")   # reversed swap
    elif len(toks) >= 3:
        # For 3 tokens: try first+last pair only (middle stripped)
        swapped = f"{toks[-1]} {toks[0]}"
        keys.append(swapped)
    return list(dict.fromkeys(keys))  # unique, preserve order


# ===========================================================================
# INVOICE LOADER — reads Phase 1 extraction Excel
# ===========================================================================

def load_invoice_data(invoice_path: str | Path) -> dict[str, dict]:
    """
    Load Phase 1 extraction Excel. Returns dict keyed by ALL candidate
    lookup keys for each employee.

    Supports dynamic column detection — works with any Phase 1 output.
    """
    path = Path(invoice_path)
    if not path.exists():
        logger.error(f"Invoice file not found: {path}")
        return {}

    xl = pd.ExcelFile(str(path))
    # Prefer sheet named 'employee', 'detail', 'data', else first sheet
    sheet = next(
        (s for s in xl.sheet_names
         if any(k in s.lower() for k in ('employee', 'detail', 'data'))),
        xl.sheet_names[0]
    )

    # Auto-detect header row
    probe = pd.read_excel(str(path), sheet_name=sheet, nrows=15, header=None)
    hrow = 0
    for i, row in probe.iterrows():
        row_str = " ".join(str(v).lower() for v in row if pd.notna(v))
        if any(kw in row_str for kw in ('name', 'plan', 'premium', 'coverage')):
            hrow = i
            break

    df = pd.read_excel(str(path), sheet_name=sheet, skiprows=hrow)
    df.columns = [str(c).strip() for c in df.columns]

    # Column detection
    col_map: dict[str, str | None] = {
        'full': None, 'first': None, 'last': None,
        'plan': None, 'premium': None, 'coverage': None,
    }
    for col in df.columns:
        cl = col.lower()
        if 'full' in cl and 'name' in cl:                  col_map['full']     = col
        elif 'employee' in cl and 'name' in cl:             col_map['full']     = col
        elif 'first' in cl and 'name' in cl:                col_map['first']    = col
        elif 'last' in cl and 'name' in cl:                 col_map['last']     = col
        elif 'plan' in cl and ('name' in cl or 'desc' in cl): col_map['plan']   = col
        elif 'premium' in cl or 'current' in cl:            col_map['premium']  = col
        elif 'coverage' in cl or 'tier' in cl:              col_map['coverage'] = col

    # Fallback: treat column 0 as full name
    if not col_map['full'] and not col_map['first']:
        col_map['full'] = df.columns[0]

    lookup: dict[str, dict] = {}
    blocked = ('total', 'subtotal', 'grand total', 'summary', 'record')

    for _, row in df.iterrows():
        # Build raw name
        if col_map['full'] and pd.notna(row.get(col_map['full'], None)):
            raw_name = str(row[col_map['full']]).strip()
        elif col_map['first'] and col_map['last']:
            f = str(row.get(col_map['first'], '') or '').strip()
            l = str(row.get(col_map['last'],  '') or '').strip()
            raw_name = f"{f} {l}".strip()
        else:
            continue

        if not raw_name or any(b in raw_name.lower() for b in blocked):
            continue

        # Clean premium
        prem_raw = row.get(col_map['premium']) if col_map['premium'] else None
        if isinstance(prem_raw, str):
            prem_raw = re.sub(r'[^\d.]', '', prem_raw)
            try:    prem_raw = float(prem_raw)
            except: prem_raw = None

        entry = {
            'raw_name': raw_name,
            'plan':     row.get(col_map['plan'])     if col_map['plan']     else None,
            'premium':  prem_raw,
            'coverage': row.get(col_map['coverage']) if col_map['coverage'] else None,
        }

        # Register under every candidate key
        for key in lookup_keys(raw_name):
            if key:
                lookup[key] = entry

    unique = len({v['raw_name'] for v in lookup.values()})
    logger.info(f"Invoice lookup built: {unique} unique employees from '{sheet}'")
    return lookup


# ===========================================================================
# FILLED EXCEL SCANNER — finds all discrepancy rows and column positions
# ===========================================================================

def _find_columns(ws) -> dict[str, int | None]:
    """
    Auto-detect key column positions from the header row of the filled Excel.
    Returns a dict: {field: column_index (1-based)} or None if not found.
    """
    cols: dict[str, int | None] = {
        'name': None, 'first': None, 'last': None,
        'plan': None, 'premium': None, 'disc': None,
    }
    for r in range(1, 40):
        row_vals = {
            c: str(ws.cell(row=r, column=c).value or '').strip().lower()
            for c in range(1, ws.max_column + 1)
        }
        joined = " ".join(row_vals.values())
        if not any(k in joined for k in ('name', 'employee', 'first')):
            continue

        for c, v in row_vals.items():
            if   ('employee' in v and 'name' in v) or ('full' in v and 'name' in v):
                cols['name']    = c
            elif 'first' in v and 'name' in v:
                cols['first']   = c
            elif 'last'  in v and 'name' in v:
                cols['last']    = c
            elif 'plan'  in v:
                cols['plan']    = c
            elif 'premium' in v:
                cols['premium'] = c
            elif 'discrep' in v:
                cols['disc']    = c

        return cols  # found header row

    return cols


def _get_name_from_row(ws, row_idx: int, cols: dict) -> str:
    """Extract employee name from a worksheet row based on detected column layout."""
    if cols['name']:
        val = ws.cell(row=row_idx, column=cols['name']).value
        return str(val).strip() if val else ""
    if cols['first'] and cols['last']:
        f = str(ws.cell(row=row_idx, column=cols['first']).value or '').strip()
        l = str(ws.cell(row=row_idx, column=cols['last']).value  or '').strip()
        return f"{f} {l}".strip()
    return ""


# ===========================================================================
# MATCHING ENGINE
# ===========================================================================

def match_name(
    raw_name: str,
    invoice_lookup: dict[str, dict],
    threshold: float = 85.0,
) -> tuple[dict | None, str, float]:
    """
    Try to find invoice_lookup entry for raw_name using 3-pass strategy.

    Returns:
        (entry, match_type, confidence)
          entry      — the matched invoice dict, or None
          match_type — 'exact' | 'token_swap' | 'fuzzy' | 'none'
          confidence — 0-100
    """
    if not raw_name:
        return None, 'none', 0.0

    # --- Pass 1 & 2: exact canonical + token-swap ---
    for key in lookup_keys(raw_name):
        if key in invoice_lookup:
            return invoice_lookup[key], 'canonical', 100.0

    # --- Pass 3: fuzzy across all known keys ---
    best_score = 0.0
    best_entry = None
    raw_canon = canonical(raw_name)

    for inv_key, entry in invoice_lookup.items():
        score = _similarity(raw_canon, inv_key)
        if score > best_score:
            best_score = score
            best_entry = entry

    if best_entry and best_score >= threshold:
        return best_entry, 'fuzzy', best_score

    if best_entry and best_score >= 50:
        return best_entry, 'possible', best_score

    return None, 'none', 0.0


# ===========================================================================
# MAIN VALIDATOR
# ===========================================================================

def run_validation(
    filled_path: str | Path,
    invoice_path: str | Path,
    output_path: str | Path | None = None,
    threshold: float = 85.0,
) -> dict:
    """
    Core validation logic — works for type1, type2, and type3 filled Excels.

    Args:
        filled_path   — Excel produced by fill_template.py
        invoice_path  — Phase 1 extraction Excel
        output_path   — destination for validated Excel (default: _validated.xlsx)
        threshold     — fuzzy match confidence threshold (0-100)

    Returns:
        dict with validation statistics and audit log entries
    """
    filled_path  = Path(filled_path)
    invoice_path = Path(invoice_path)

    if output_path is None:
        output_path = filled_path.with_stem(filled_path.stem + "_validated")
    output_path = Path(output_path)

    if not filled_path.exists():
        logger.error(f"Filled Excel not found: {filled_path}")
        return {}
    if not invoice_path.exists():
        logger.error(f"Invoice Excel not found: {invoice_path}")
        return {}

    # Load invoice lookup
    invoice_lookup = load_invoice_data(invoice_path)
    if not invoice_lookup:
        logger.error("Invoice lookup is empty — cannot validate.")
        return {}

    # Open filled workbook
    wb = load_workbook(str(filled_path))
    ws = next(
        (wb[s] for s in wb.sheetnames
         if any(k in s.lower() for k in ('census', 'employee', 'table', 'sheet'))),
        wb.active
    )

    col_positions = _find_columns(ws)
    disc_col = col_positions.get('disc')
    plan_col = col_positions.get('plan')
    prem_col = col_positions.get('premium')

    if disc_col is None:
        logger.error("Discrepancies column not found in filled Excel.")
        return {}

    # Find data start row (row after header)
    data_start = 2
    for r in range(1, 40):
        row_vals = {
            c: str(ws.cell(row=r, column=c).value or '').strip().lower()
            for c in range(1, ws.max_column + 1)
        }
        joined = " ".join(row_vals.values())
        if any(k in joined for k in ('name', 'employee', 'first')):
            data_start = r + 1
            break

    # ------------------------------------------------------------------
    # Scan all rows for discrepancy flags
    # ------------------------------------------------------------------
    audit_log: list[dict] = []
    stats = {
        'total_rows':         0,
        'already_correct':    0,
        'resolved_canonical': 0,
        'resolved_fuzzy':     0,
        'still_unresolved':   0,
        'possible_matches':   0,
        'appended_deleted':   0,
    }

    claimed_invoices = set()
    rows_to_delete = []

    for row_idx in range(data_start, ws.max_row + 1):
        # Stop at first completely empty row
        row_is_empty = all(
            ws.cell(row=row_idx, column=c).value is None
            for c in range(1, min(ws.max_column + 1, 6))
        )
        if row_is_empty:
            break

        disc_cell = ws.cell(row=row_idx, column=disc_col)
        disc_val  = str(disc_cell.value or '').strip()
        raw_name  = _get_name_from_row(ws, row_idx, col_positions)

        if not raw_name and not disc_val:
            continue

        stats['total_rows'] += 1

        # ── Rows that need resolution ─────────────────────────────────
        is_not_census  = _NOT_ON_CENSUS.lower()  in disc_val.lower()
        is_not_invoice = _NOT_ON_INVOICE.lower() in disc_val.lower()

        # ── Already resolved rows — skip ──────────────────────────────
        if not (is_not_census or is_not_invoice):
            stats['already_correct'] += 1
            if 'matched' in disc_val.lower():
                # If we run on an already validated file, track previously claimed matches
                pass
            continue

        # Run name matching
        match_entry, match_type, confidence = match_name(
            raw_name, invoice_lookup, threshold
        )

        audit_entry = {
            'row':            row_idx,
            'raw_name':       raw_name,
            'original_status': disc_val,
            'match_type':     match_type,
            'confidence':     round(confidence, 1),
            'matched_to':     match_entry['raw_name'] if match_entry else None,
            'action':         'unresolved'
        }

        if is_not_invoice:
            # ── Original Census Row needing invoice data ──────────────────
            if match_type == 'canonical':
                _apply_match(ws, row_idx, match_entry, col_positions,
                             f"Matched -> {match_entry['raw_name']}"[:40],
                             _FILL_GREEN)
                stats['resolved_canonical'] += 1
                audit_entry['action'] = 'resolved_canonical'
                claimed_invoices.add(match_entry['raw_name'])

            elif match_type == 'fuzzy' and confidence >= threshold:
                label = f"Fuzzy Match ({confidence:.0f}%) -> {match_entry['raw_name']}"[:40]
                _apply_match(ws, row_idx, match_entry, col_positions, label, _FILL_YELLOW)
                stats['resolved_fuzzy'] += 1
                audit_entry['action'] = 'resolved_fuzzy'
                claimed_invoices.add(match_entry['raw_name'])

            elif match_type == 'possible':
                disc_cell.value      = f"Possible Match ({confidence:.0f}%) -> {match_entry['raw_name']}"[:40]
                disc_cell.fill       = _FILL_ORANGE
                disc_cell.font       = _FONT
                disc_cell.alignment  = _CENTER
                stats['possible_matches'] += 1
                audit_entry['action'] = 'flagged_possible'

            else:
                disc_cell.fill      = _FILL_RED
                disc_cell.font      = _FONT
                disc_cell.alignment = _CENTER
                stats['still_unresolved'] += 1
                audit_entry['action'] = 'unresolved'

        elif is_not_census:
            # ── Appended Invoice Row (Added by Phase 2) ───────────────────
            if match_entry and match_entry['raw_name'] in claimed_invoices:
                # We already mapped this invoice record to a proper census row above!
                rows_to_delete.append(row_idx)
                stats['appended_deleted'] += 1
                audit_entry['action'] = 'deleted_duplicate'
            else:
                # Truly not on the original census, keeping it as an appended exception.
                disc_cell.value      = _NOT_ON_CENSUS
                disc_cell.fill       = _FILL_RED
                disc_cell.font       = _FONT
                disc_cell.alignment  = _CENTER
                stats['still_unresolved'] += 1
                audit_entry['action'] = 'kept_unresolved_appended'

        audit_log.append(audit_entry)

    # ------------------------------------------------------------------
    # Delete duplicate appended rows (safely backwards)
    # ------------------------------------------------------------------
    for row_idx in reversed(rows_to_delete):
        ws.delete_rows(row_idx, amount=1)

    # ------------------------------------------------------------------
    # Save validated workbook
    # ------------------------------------------------------------------
    wb.save(str(output_path))
    logger.info(f"Validated Excel saved -> {output_path}")

    # ------------------------------------------------------------------
    # Save audit log JSON
    # ------------------------------------------------------------------
    audit_path = output_path.with_suffix('.audit.json')
    with open(str(audit_path), 'w', encoding='utf-8') as f:
        json.dump({'stats': stats, 'entries': audit_log}, f, indent=2, default=str)
    logger.info(f"Audit log saved -> {audit_path}")

    # Summary
    logger.info(
        "\n" + "="*55 + "\n"
        "  VALIDATION SUMMARY\n"
        f"  Total rows scanned   : {stats['total_rows']}\n"
        f"  Already correct      : {stats['already_correct']}\n"
        f"  Resolved (canonical) : {stats['resolved_canonical']}\n"
        f"  Resolved (fuzzy)     : {stats['resolved_fuzzy']}\n"
        f"  Possible matches     : {stats['possible_matches']} (review needed)\n"
        f"  Still unresolved     : {stats['still_unresolved']}\n"
        f"  Appended deleted     : {stats['appended_deleted']} (removed duplicates)\n"
        + "="*55
    )

    return {
        'stats':       stats,
        'output_path': str(output_path),
        'audit_path':  str(audit_path),
        'audit_log':   audit_log,
    }


def _apply_match(ws, row_idx: int, entry: dict, cols: dict,
                 label: str, fill: PatternFill) -> None:
    """Write matched invoice data into the worksheet row and update Discrepancy cell."""
    disc_col = cols.get('disc')
    plan_col = cols.get('plan')
    prem_col = cols.get('premium')

    # Update Discrepancy cell
    if disc_col:
        cell            = ws.cell(row=row_idx, column=disc_col)
        cell.value      = label
        cell.fill       = fill
        cell.font       = _FONT
        cell.alignment  = _CENTER

    # Fill plan name if cell is empty
    if plan_col and entry.get('plan'):
        plan_cell = ws.cell(row=row_idx, column=plan_col)
        if not plan_cell.value:
            plan_cell.value     = entry['plan']
            plan_cell.font      = _FONT
            plan_cell.alignment = _LEFT

    # Fill premium if cell is empty
    if prem_col and entry.get('premium') is not None:
        prem_cell = ws.cell(row=row_idx, column=prem_col)
        if not prem_cell.value:
            prem_cell.value         = entry['premium']
            prem_cell.font          = _FONT
            prem_cell.alignment     = _CENTER
            prem_cell.number_format = '$#,##0.00'


# ===========================================================================
# CLI ENTRY POINT
# ===========================================================================

def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Post-fill Data Validator — resolves name mismatches between "
            "census and invoice using fuzzy name normalisation. "
            "Works for type1, type2, and type3 filled Excels."
        )
    )
    parser.add_argument(
        'filled_excel',
        help='Path to the filled census Excel (output of fill_template.py)'
    )
    parser.add_argument(
        'invoice_excel',
        help='Path to the Phase 1 invoice extraction Excel'
    )
    parser.add_argument(
        'output', nargs='?', default=None,
        help='Output path for validated Excel (default: <filled>_validated.xlsx)'
    )
    parser.add_argument(
        '--threshold', type=float, default=85.0,
        help='Fuzzy match minimum confidence 0-100 (default: 85)'
    )
    args = parser.parse_args()

    result = run_validation(
        filled_path   = args.filled_excel,
        invoice_path  = args.invoice_excel,
        output_path   = args.output,
        threshold     = args.threshold,
    )

    if not result:
        return 1

    try:
        print(f"\n[OK] Validated Excel : {result['output_path']}")
        print(f"[OK] Audit log       : {result['audit_path']}")
        print(f"[OK] Stats           : {result['stats']}")
    except Exception:
        pass  # Suppress any encoding errors on Windows cp1252 terminals
    return 0


if __name__ == "__main__":
    sys.exit(main())
