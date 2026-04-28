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
    Strategy: extract first+last tokens after stripping suffixes,
    then check both same-order and reversed-order matches.
    """
    t1 = _name_tokens(extracted_name)
    t2 = _name_tokens(invoice_name)

    if not t1 or not t2:
        return False

    # Full normalised string match (after suffix stripping)
    if " ".join(t1) == " ".join(t2):
        return True

    # Both have at least 2 meaningful tokens — compare first+last in both orderings
    if len(t1) >= 2 and len(t2) >= 2:
        # Same order: t1[0]==t2[0] and t1[-1]==t2[-1]
        if t1[0] == t2[0] and t1[-1] == t2[-1]:
            return True
        # Reversed order: t1[0]==t2[-1] and t1[-1]==t2[0]  (LAST FIRST vs First Last)
        if t1[0] == t2[-1] and t1[-1] == t2[0]:
            return True

    # Single-token fallback
    if len(t1) == 1 and len(t2) == 1:
        return t1[0] == t2[0]

    return False


def coverage_match(extracted_coverage: Any, invoice_coverage: Any) -> bool:
    """
    Match coverage tiers. If invoice coverage is absent/empty,
    treat as unverifiable and return True (avoid false mismatches).
    """
    inv_tier = canonical_coverage_tier(invoice_coverage)
    if not inv_tier:
        return True
    return canonical_coverage_tier(extracted_coverage) == inv_tier


def canonical_coverage_tier(value: Any) -> str:
    """Normalize common coverage-tier aliases across census and invoice files.
    Handles short codes (EE, ES, EC, FAM) AND long-form strings from Curative-style invoices:
      'Employee Only'                          → EE
      'Employee and 1+ Children'               → EC
      'Employee, Spouse, and 1+ Children'      → FAM
      'Employee and Spouse'                    → ES
    Falls back to keyword scanning for any unrecognised long-form string.
    """
    token = normalize_text(value).replace(" ", "").upper()

    tier_map = {
        # Short codes
        "E":                         "EE",
        "EE":                        "EE",
        "S":                         "EE",
        "EMPLOYEE":                  "EE",
        "EMPLOYER":                  "EE",
        "EMPLOYEEONLY":              "EE",
        # Employee + Children
        "EC":                        "EC",
        "CH":                        "EC",
        "SD":                        "EC",
        "EMPLOYEECHILDREN":          "EC",
        "EMPLOYEECHILD(REN)":        "EC",
        "EMPLOYEEAND1CHILDREN":      "EC",
        "EMPLOYEEAND1+CHILDREN":     "EC",
        "EMPLOYEEANDCHILDREN":       "EC",
        "EMPLOYEEAND1":              "EC",   # truncated form
        "EMPLOYEEAND1+":             "EC",
        # Employee + Spouse
        "ES":                        "ES",
        "SP":                        "ES",
        "SPOUSE":                    "ES",
        "SS":                        "ES",
        "EMPLOYEESPOUSE":            "ES",
        "EMPLOYEEANDSPOUSE":         "ES",
        "EMPLOYEESPOUSEONLY":        "ES",
        # Family
        "EF":                        "FAM",
        "F":                         "FAM",
        "FAM":                       "FAM",
        "FAMILY":                    "FAM",
        "EMPLOYEEFAMILY":            "FAM",
        "EMPLOYEEANDSPOUSEANDCHILD": "FAM",
        "EMPLOYEESPOUSEANDCHILDREN": "FAM",
        "EMPLOYEESPOUSEAND1+CHILDREN": "FAM",
        "EMPLOYEESPOUSEAND1CHILDREN":  "FAM",
        "SPOUSEEMPLOYEE":            "FAM",   # reversed token artefact
        "SPOUSEEMPLOYEE1+":          "FAM",
    }

    result = tier_map.get(token)
    if result:
        return result

    # ── Keyword fallback for unrecognised long-form strings ───────────────────
    # e.g. "Employee, Spouse, and 1+ Children" / "Employee and 1+"
    raw_lower = str(value or '').lower()
    has_spouse   = 'spouse' in raw_lower or 'partner' in raw_lower
    has_children = 'child' in raw_lower or '1+' in raw_lower or 'dependent' in raw_lower

    if has_spouse and has_children:
        return "FAM"
    if has_spouse:
        return "ES"
    if has_children:
        return "EC"
    if 'only' in raw_lower or 'employee' in raw_lower:
        return "EE"

    return token  # return the normalised token as-is if nothing matches


def discrepancy_status(
    extracted_name: Any,
    invoice_name: Any,
    extracted_coverage_tier: Any,
    invoice_coverage_tier: Any,
) -> str:
    """
    Return validation status for Discrepancies column.
      Correct                                    – name and coverage both match
      mismatch employee name                     – name differs
      mismatch coverage name                     – coverage differs
      mismatch employee name & mismatch coverage name
    """
    n_match = names_match(extracted_name, invoice_name)
    c_match = coverage_match(extracted_coverage_tier, invoice_coverage_tier)

    if n_match and c_match:
        return "Correct"

    reasons = []
    if not n_match:
        reasons.append("mismatch employee name")
    if not c_match:
        reasons.append("mismatch coverage name")

    return " & ".join(reasons)