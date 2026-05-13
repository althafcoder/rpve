import re
from typing import Any

import pandas as pd

NOT_ON_CENSUS_STATUS = "Not on census"


def normalize_text(value: Any) -> str:
    """Normalize text for strict-but-stable equality comparison."""
    if value is None or pd.isna(value):
        return ""

    text = str(value).strip().lower()
    # Handle "Last, First" format by flipping it
    if ',' in text:
        parts = [p.strip() for p in text.split(',')]
        if len(parts) >= 2:
            text = f"{parts[1]} {parts[0]}"

    text = re.sub(r"[^a-z0-9\s]", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def names_match(extracted_name: Any, invoice_name: Any) -> bool:
    n1 = normalize_text(extracted_name)
    n2 = normalize_text(invoice_name)
    if n1 == n2:
        return True

    # Check first and last name match to handle middle name/initial differences
    p1, p2 = n1.split(), n2.split()
    if len(p1) >= 2 and len(p2) >= 2:
        return p1[0] == p2[0] and p1[-1] == p2[-1]

    return False


def coverage_match(extracted_coverage: Any, invoice_coverage: Any) -> bool:
    return canonical_coverage_tier(extracted_coverage) == canonical_coverage_tier(
        invoice_coverage
    )


def canonical_coverage_tier(value: Any) -> str:
    """Normalize common coverage-tier aliases across census and invoice files."""
    token = normalize_text(value).replace(" ", "").upper()

    tier_map = {
        "E": "EE",
        "EE": "EE",
        "S": "EE",
        "EMPLOYEE": "EE",
        "EMPLOYER": "EE",
        "EC": "EC",
        "CH": "EC",
        "SD": "EC",
        "EMPLOYEECHILDREN": "EC",
        "ES": "ES",
        "SP": "ES",
        "SPOUSE": "ES",
        "SS": "ES",
        "EMPLOYEESPOUSE": "ES",
        "EF": "FAM",
        "F": "FAM",
        "FAM": "FAM",
        "FAMILY": "FAM",
        "EMPLOYEEFAMILY": "FAM",
    }

    return tier_map.get(token, token)


def discrepancy_status(
    extracted_name: Any,
    invoice_name: Any,
    extracted_coverage_tier: Any,
    invoice_coverage_tier: Any,
    name_is_matched: bool = False
) -> str:
    """
    Return validation status for Discrepancies column.
    - Correct: employee name and coverage tier both match
    - Specific mismatch messages otherwise
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
