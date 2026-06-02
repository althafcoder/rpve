"""
pdf_extractor.py  —  Technology 1: pdfplumber
==============================================
Extracts employee / premium data from a California Choice PDF invoice.

Usage (standalone):
    python pdf_extractor.py "input/CA Choice 3_....pdf"

The PDF has a tabular section per employee like:
    Bradvica,Andrew  Medical  Employee  KaiserPermanente P PHA  $1249.33  ...
    ...              Dental   Waived
    ...              Vision   Waived

This module parses those rows and returns a structured list:
    [
      {
        'raw_name':  'Bradvica,Andrew',
        'first':     'Andrew',
        'last':      'Bradvica',
        'coverage':  'EE',          # from coverage tier column
        'plan':      'Kaiser Permanente P PHA',
        'premium':   1249.33,
      },
      ...
    ]
"""

import re
import sys
import json
import logging
from pathlib import Path

import pdfplumber

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Coverage tier mapping for CA Choice labels
# ---------------------------------------------------------------------------
_COVERAGE_MAP = {
    "employee":              "EE",
    "ee":                    "EE",
    "ee+family":             "FAM",
    "ee + family":           "FAM",
    "ee+1":                  "ES",
    "ee + 1":                "ES",
    "ee+spouse":             "ES",
    "ee + spouse":           "ES",
    "ee+children":           "EC",
    "ee + children":         "EC",
    "employee+family":       "FAM",
    "employee + family":     "FAM",
    "employee+1":            "ES",
    "employee + 1":          "ES",
    "employee+spouse":       "ES",
    "employee + spouse":     "ES",
    "employee+children":     "EC",
    "employee + children":   "EC",
    "ee+family":             "FAM",
    "eefamily":              "FAM",
    "eespouse":              "ES",
    "eechildren":            "EC",
    "waived":                "WO",
    "waiver":                "WO",
    "cobra":                 "C",
}


def _normalize_coverage(raw: str) -> str:
    """Map CA Choice coverage string to canonical tier code."""
    key = re.sub(r"\s+", " ", str(raw or "")).strip().lower()
    if key in _COVERAGE_MAP:
        return _COVERAGE_MAP[key]
    # Try removing spaces
    compact = key.replace(" ", "")
    if compact in _COVERAGE_MAP:
        return _COVERAGE_MAP[compact]
    return key.upper() if key else "EE"


def _parse_name(raw_name: str):
    """
    Split 'Last,First' or 'Last First' or 'VanWonterghem,Mark' into
    (first, last).  Returns (first, last).
    """
    raw = raw_name.strip()
    if "," in raw:
        parts = [p.strip() for p in raw.split(",", 1)]
        last  = parts[0].strip()
        first = parts[1].strip() if len(parts) > 1 else ""
    else:
        # Heuristic: last word is last name? Not reliable for compound surnames.
        # For CA Choice PDFs the format is nearly always "Last,First" so fallback:
        tokens = raw.split()
        if len(tokens) >= 2:
            last  = tokens[0]
            first = " ".join(tokens[1:])
        else:
            last  = raw
            first = ""
    return first, last


def _clean_dollar(value: str):
    """'$ 1,249.33' → 1249.33 (float) or None."""
    s = re.sub(r"[^\d.]", "", str(value or ""))
    try:
        return float(s) if s else None
    except ValueError:
        return None


# Known CA Choice carrier/plan name fragments that get merged in PDF text
_PLAN_NAME_FIXES = [
    (r"KaiserPermanente",      "Kaiser Permanente"),
    (r"AnthemBlueCross",       "Anthem Blue Cross"),
    (r"SharpHealthPlan",       "Sharp Health Plan"),
    (r"UnitedHealthCare",      "UnitedHealth Care"),
    (r"BlueCross",             "Blue Cross"),
    (r"BlueShield",            "Blue Shield"),
    (r"HealthNet",             "Health Net"),
]


def _fix_plan_name(name: str) -> str:
    """Restore spaces in concatenated plan names from PDF extraction."""
    for pattern, replacement in _PLAN_NAME_FIXES:
        name = re.sub(pattern, replacement, name, flags=re.IGNORECASE)
    return name.strip()


# ---------------------------------------------------------------------------
# Main extraction logic
# ---------------------------------------------------------------------------

# Lines that are page headers / boiler-plate — not employee rows
_SKIP_PATTERNS = [
    r"^california\s*choice",
    r"^invoice",
    r"^b2r\s*consulting",
    r"^invoicenumber",
    r"^groupnumber",
    r"^duedate",
    r"^coverageperiod",
    r"^premium\s*payment",
    r"^check\s*your\s*next",
    r"^medical\s*tier",
    r"^employer\s*contribution",
    r"^employee\s+plan\s+coverage",  # table header row
    r"^information\s+type",          # table header continuation
    r"^summary\s*of",
    r"^plan\s+type",
    r"^medical\s+anthem",
    r"^medical\s+kaiser",
    r"^medical\s+sharp",
    r"^dental\s+ameritas",
    r"^dental\s+smile",
    r"^vision\s+eye",
    r"^mandated",
    r"^your\s+health",
    r"^listed\s+in",
    r"^coverage\s+can",
    r"^about\s+your",
    r"^note:",
    r"^page\s*\d",
    r"^\d+$",                        # lone page number
]
_SKIP_RE = [re.compile(p, re.IGNORECASE) for p in _SKIP_PATTERNS]


def _is_skip_line(text: str) -> bool:
    t = text.strip()
    if not t:
        return True
    for pat in _SKIP_RE:
        if pat.match(t):
            return True
    return False


# Employee line: starts with "Last,First" (or similar) followed by
# "Medical" or "Dental" or "Vision"  (the benefit type column).
# Pattern allows for attached digits like "Bradvica,Andrew2078"
_EMPLOYEE_LINE_RE = re.compile(
    r"^([A-Za-z][A-Za-z\-]+(?:\s+[A-Za-z]+)*,\s*[A-Za-z][A-Za-z\-]+(?:\s+[A-Za-z]+)*)"
    r"(?:\d{1,6})?"                    # optional member/ID digits glued to name
    r"\s+(Medical|Dental|Vision)"
    r"\s+(.+)",
    re.IGNORECASE
)

# Continuation lines for Dental / Vision benefit under the same employee
_BENEFIT_LINE_RE = re.compile(
    r"^\s*(Dental|Vision)\s+(.+)",
    re.IGNORECASE
)

# Dollar amount anywhere on a line
_DOLLAR_RE = re.compile(r"\$\s*([\d,]+\.\d{2})")


def _extract_medical_info(rest: str):
    """
    Parse the columns after 'Medical  Employee  KaiserPermanente P PHA  $ 1249.33 ...'
    Returns (coverage_raw, plan_name, total_premium).
    rest is everything after 'Medical' on the same line.
    """
    # Strip trailing dollar amounts to get plan info
    dollars = _DOLLAR_RE.findall(rest)
    # The FIRST dollar amount after the employee/tier section is the Employee Premium,
    # and the THIRD is the Total (EE+Dep).  We want col "Employee Plan Premium" which
    # is the FIRST or THIRD — for medical verification we use total premium col.
    total_premium = None
    if dollars:
        # Clean & parse all dollar values
        amounts = []
        for d in dollars:
            try:
                amounts.append(float(d.replace(",", "")))
            except ValueError:
                pass
        # Column layout: Employee Premium | Dependent Premium | Employer Total
        # | Employer Contribution | Employee Contribution | EE Deduction
        # We use the THIRD value (Employer Total) as "total monthly premium" for the record
        if len(amounts) >= 3:
            total_premium = amounts[2]
        elif amounts:
            total_premium = amounts[0]

    # Remove dollar fields to extract coverage + plan
    stripped = re.sub(r"\$\s*[\d,]+\.\d{2}", " ", rest)
    stripped = re.sub(r"\s+", " ", stripped).strip()

    # Tokenize
    tokens = stripped.split()
    if not tokens:
        return "EE", "", total_premium

    coverage_raw = tokens[0]  # e.g. "Employee", "EE+Family", "Waived"
    plan_tokens  = tokens[1:] if len(tokens) > 1 else []

    # Keep plan tier codes (PHA, SHD, etc.) — they identify the exact carrier tier
    plan_name = " ".join(plan_tokens).strip()

    # Strip trailing 2-char charge code (e.g. "CA" at end of some lines)
    plan_name = re.sub(r"\s+[A-Z]{2}\s*$", "", plan_name).strip()

    # Fix concatenated words from PDF text extraction
    plan_name = _fix_plan_name(plan_name)

    # If coverage is Waived, there is no real plan or premium
    cov_lower = coverage_raw.strip().lower()
    if cov_lower in ("waived", "waiver"):
        return coverage_raw, "", None

    return coverage_raw, plan_name, total_premium


def extract_employees_from_pdf(pdf_path: str) -> list:
    """
    Parse a California Choice PDF invoice and return a list of employee dicts.
    Each dict has keys: raw_name, first, last, coverage, plan, premium.
    Only Medical rows are used for plan/premium (Dental/Vision are ignored for
    the RAPT census fill, which only captures the primary medical plan).
    """
    employees = []
    current_emp = None

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            lines = text.splitlines()

            for raw_line in lines:
                line = raw_line.strip()
                # Collapse run-together words (PDF extraction artefact)
                # e.g. "KaiserPermanente" is already glued — keep as-is

                if _is_skip_line(line):
                    continue

                m = _EMPLOYEE_LINE_RE.match(line)
                if m:
                    # Save previous employee if any
                    if current_emp and current_emp.get("_has_medical"):
                        employees.append(current_emp)

                    raw_name   = m.group(1).strip()
                    benefit    = m.group(2).strip()       # Medical / Dental / Vision
                    rest       = m.group(3).strip()

                    first, last = _parse_name(raw_name)

                    current_emp = {
                        "raw_name": raw_name,
                        "first":    first,
                        "last":     last,
                        "coverage": "EE",
                        "plan":     "",
                        "premium":  None,
                        "_has_medical": False,
                    }

                    if benefit.lower() == "medical":
                        cov_raw, plan_name, total_premium = _extract_medical_info(rest)
                        current_emp["coverage"] = _normalize_coverage(cov_raw)
                        current_emp["plan"]     = plan_name
                        current_emp["premium"]  = total_premium
                        current_emp["_has_medical"] = True

                    continue

                # Continuation: Dental / Vision line under current employee
                if current_emp:
                    bm = _BENEFIT_LINE_RE.match(line)
                    if bm:
                        # We don't currently capture dental/vision into the census
                        continue

                    # Lines like "$ 624.66" or "$ 289.78 CA" are employee totals — skip
                    if re.match(r"^\$\s*[\d,]+\.\d{2}", line):
                        continue

        # Don't forget the last employee
        if current_emp and current_emp.get("_has_medical"):
            employees.append(current_emp)

    # Remove internal helper key
    for emp in employees:
        emp.pop("_has_medical", None)

    logger.info(f"Extracted {len(employees)} medical employees from PDF '{pdf_path}'")
    for e in employees:
        logger.debug(f"  {e['last']}, {e['first']}  coverage={e['coverage']}  "
                     f"plan={e['plan']}  premium={e['premium']}")

    return employees


# ---------------------------------------------------------------------------
# CLI entry-point (standalone test)
# ---------------------------------------------------------------------------
def main():
    if len(sys.argv) < 2:
        print(f"Usage: python {Path(__file__).name} <invoice.pdf>")
        sys.exit(1)

    pdf_path = sys.argv[1]
    employees = extract_employees_from_pdf(pdf_path)

    print(f"\n{'='*60}")
    print(f"Extracted {len(employees)} employees from: {pdf_path}")
    print(f"{'='*60}")
    for emp in employees:
        print(
            f"  {emp['last']:20s} {emp['first']:15s}  "
            f"Coverage: {emp['coverage']:5s}  "
            f"Plan: {emp['plan']:40s}  "
            f"Premium: {emp['premium']}"
        )


if __name__ == "__main__":
    main()
