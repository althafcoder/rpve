"""
validation.py — Discrepancy status helpers for type4
=====================================================
Shared validation logic: name matching, coverage-tier normalisation,
and the final discrepancy_status() used by fill_template.py.
"""
import re
from typing import Any

import pandas as pd

NOT_ON_CENSUS_STATUS = "Not on census"


def normalize_text(value: Any) -> str:
    """Normalize text for comparison — strips punctuation, extra spaces, lowercases."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).strip().lower()
    # Handle "Last, First" comma format
    if ',' in text:
        parts = [p.strip() for p in text.split(',')]
        if len(parts) >= 2:
            text = f"{parts[1]} {parts[0]}"
    text = re.sub(r"[^a-z0-9\s]", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def _name_tokens(value: Any) -> list:
    """Return cleaned word tokens for a name, stripping suffixes like jr/sr/ii/iii."""
    SUFFIXES = {'jr', 'sr', 'ii', 'iii', 'iv', 'v', 'f', 'esq'}
    raw = normalize_text(value)
    return [t for t in raw.split() if t not in SUFFIXES]


def names_match(extracted_name: Any, invoice_name: Any) -> bool:
    """
    Match names regardless of:
      - Case differences
      - LAST FIRST vs First Last ordering  (invoice uses LAST FIRST all-caps)
      - Middle names / initials / suffixes (jr, sr, f, etc.)
    """
    t1 = _name_tokens(extracted_name)
    t2 = _name_tokens(invoice_name)

    if not t1 or not t2:
        return False

    if " ".join(t1) == " ".join(t2):
        return True

    if len(t1) >= 2 and len(t2) >= 2:
        if t1[0] == t2[0] and t1[-1] == t2[-1]:
            return True
        if t1[0] == t2[-1] and t1[-1] == t2[0]:
            return True

    if len(t1) == 1 and len(t2) == 1:
        return t1[0] == t2[0]

    return False


def canonical_coverage_tier(value: Any) -> str:
    """Normalize common coverage-tier aliases across census and invoice files."""
    token = normalize_text(value).replace(" ", "").upper()

    tier_map = {
        "E": "EE", "EE": "EE",
        "S": "ES", "C": "EC", "F": "FAM", "W": "WO",
        "NC": "RC", "NE": "NE",
        "EMPLOYEE": "EE", "EMPLOYER": "EE", "EMPLOYEEONLY": "EE",
        "EC": "EC", "CH": "EC", "SD": "EC",
        "EMPLOYEECHILDREN": "EC", "EMPLOYEEANDCHILDREN": "EC",
        "EMPLOYEEAND1CHILDREN": "EC", "EMPLOYEEAND1+CHILDREN": "EC",
        "ES": "ES", "SP": "ES", "SPOUSE": "ES", "SS": "ES",
        "EMPLOYEESPOUSE": "ES", "EMPLOYEEANDSPOUSE": "ES",
        "EF": "FAM", "FAM": "FAM", "FAMILY": "FAM",
        "EMPLOYEEFAMILY": "FAM", "EMPLOYEEANDSPOUSEANDCHILD": "FAM",
        "EMPLOYEESPOUSEANDCHILDREN": "FAM",
        # California Choice specific
        "EMPLOYEE": "EE",
        "EE+FAMILY": "FAM", "EE+SPOUSE": "ES", "EE+CHILDREN": "EC",
        "EEFAMILY": "FAM", "EESPOUSE": "ES", "EECHILDREN": "EC",
    }

    result = tier_map.get(token)
    if result:
        return result

    raw_lower = str(value or '').lower()
    has_spouse   = 'spouse' in raw_lower or 'partner' in raw_lower
    has_children = 'child' in raw_lower or '1+' in raw_lower or 'dependent' in raw_lower or 'family' in raw_lower

    if has_spouse and has_children:
        return "FAM"
    if has_spouse:
        return "ES"
    if has_children:
        return "EC"
    if 'only' in raw_lower or 'employee' in raw_lower:
        return "EE"

    return token


def coverage_match(extracted_coverage: Any, invoice_coverage: Any) -> bool:
    """
    Match coverage tiers. If invoice coverage is absent/empty,
    treat as unverifiable and return True (avoid false mismatches).
    """
    inv_tier = canonical_coverage_tier(invoice_coverage)
    if not inv_tier:
        return True
    return canonical_coverage_tier(extracted_coverage) == inv_tier


def discrepancy_status(
    extracted_name: Any,
    invoice_name: Any,
    extracted_coverage_tier: Any,
    invoice_coverage_tier: Any,
    name_is_matched: bool = False
) -> str:
    """
    Return validation status for Discrepancies column.
      'Correct'                                    – name and coverage both match
      'mismatch employee name'                     – name differs
      'mismatch coverage name'                     – coverage differs
      'mismatch employee name & mismatch coverage name'
    """
    n_match = name_is_matched or names_match(extracted_name, invoice_name)
    c_match = coverage_match(extracted_coverage_tier, invoice_coverage_tier)

    if n_match and c_match:
        return "Correct"

    reasons = []
    if not n_match:
        reasons.append("mismatch employee name")
    if not c_match:
        reasons.append("mismatch coverage name")

    return " & ".join(reasons)
