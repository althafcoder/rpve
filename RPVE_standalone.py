"""
RPVE_standalone.py
==================
Standalone FastAPI server for the RPVE Benefit Invoice Extractor.

RPVE = Resourcing · Prestige · Velocity · Engage

USAGE
-----
1. pip install -r requirements_RPVE.txt
2. Add OPENAI_API_KEY=sk-... to .env
3. python RPVE_standalone.py
4. Open http://localhost:8009

ENDPOINTS
---------
POST /extract          Upload PDF -> JSON + Excel download link
GET  /download/{file}  Download generated Excel
GET  /health           Health check
GET  /                 Serves RPVE_ui.html
"""

import os, re, json, shutil
import pdfplumber
from pathlib import Path
from datetime import datetime
from dotenv import load_dotenv
from fastapi import FastAPI, File, UploadFile, HTTPException, Form
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from openai import OpenAI
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

load_dotenv()

BASE_DIR   = Path(__file__).parent
UPLOAD_DIR = BASE_DIR / "rpve_uploads"
OUTPUT_DIR = BASE_DIR / "rpve_outputs"
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

# ══════════════════════════════════════════════════════════════════════════════
# EXACT FIELD SCHEMAS - only required fields, exactly as per the SVG diagram
# ══════════════════════════════════════════════════════════════════════════════

EMPLOYEE_FIELDS = {

    # ── Engage: 5 fields ────────────────────────────────────────────────
    "engage": [
        "FIRST_NAME",
        "LAST_NAME",
        "COVERAGE",
        "PLAN_NAME",
        "COVERAGE_OPTION",
        "CURRENT_PREMIUM",
    ],

    # ── Velocity: 5 fields ──────────────────────────────────────────────
    "velocity": [
        "FIRST_NAME",
        "LAST_NAME",
        "COVERAGE",
        "PLAN_NAME",
        "COVERAGE_OPTION",
        "CURRENT_PREMIUM",
    ],

    # ── Resourcing (Kaiser + UHC): 10 required fields ────────────────────────
    # Current Plan Enrolled and Monthly Total Premium are NOT included (not required)
    "resourcing_kaiser": [
        "DATA_ROW",
        "FIRST_NAME",
        "LAST_NAME",
        "GENDER",
        "DATE_OF_BIRTH",
        "HOME_ZIP_CODE",
        "RELATIONSHIP_TO_EMPLOYEE",
        "DEPENDENT_OF_EMPLOYEE_ROW",
        "COVERAGE",
        "COBRA_PARTICIPANT",
        "CURRENT_PLAN_ENROLLED",
        "MONTHLY_TOTAL_PREMIUM",
    ],
    "resourcing_uhc": [
        "DATA_ROW",
        "FIRST_NAME",
        "LAST_NAME",
        "GENDER",
        "DATE_OF_BIRTH",
        "HOME_ZIP_CODE",
        "RELATIONSHIP_TO_EMPLOYEE",
        "DEPENDENT_OF_EMPLOYEE_ROW",
        "COVERAGE",
        "COBRA_PARTICIPANT",
        "CURRENT_PLAN_ENROLLED",
        "MONTHLY_TOTAL_PREMIUM",
    ],

    # ── Prestige (Aetna): 9 required fields ─────────────────────────────────
    # Date of Birth and Home Zip Code are NOT included (not required)
    "prestige": [
        "DATA_ROW",
        "FIRST_NAME",
        "LAST_NAME",
        "GENDER",
        "RELATIONSHIP_TO_EMPLOYEE",
        "DEPENDENT_OF_EMPLOYEE_ROW",
        "COVERAGE",
        "COBRA_PARTICIPANT",
        "PLAN_NAME",
    ],
}

# EV_OFF Strict Requirement: All 14 fields
EV_OFF_FIELDS = [
    "DATA_ROW",
    "FIRST_NAME",
    "LAST_NAME",
    "GENDER",
    "DATE_OF_BIRTH",
    "HOME_ZIP_CODE",
    "RELATIONSHIP_TO_EMPLOYEE",
    "DEPENDENT_OF_EMPLOYEE_ROW",
    "COVERAGE",
    "COVERAGE_OPTION",
    "COBRA_PARTICIPANT",
    "CURRENT_PLAN_ENROLLED",
    "PLAN_NAME",
    "MONTHLY_TOTAL_PREMIUM",
]

# EV_ON Strict Requirement: 6 fields
EV_ON_FIELDS = [
    "FIRST_NAME",
    "LAST_NAME",
    "COVERAGE",
    "PLAN_NAME",
    "COVERAGE_OPTION",
    "CURRENT_PREMIUM",
]

SUMMARY_FIELDS = {
    "engage":            ["COMPANY_NAME", "INVOICE_NUMBER", "BILLING_DATE", "DUE_DATE", "REFERENCE_NUMBER", "TOTAL_COST"],
    "velocity":          ["COMPANY_NAME", "CLIENT_NUMBER", "STATEMENT_DATE", "STATEMENT_NUMBER", "GRAND_TOTAL"],
    "prestige":          ["COMPANY_NAME", "BILL_NUMBER", "TRIAD_NUMBER", "ACCOUNT_NUMBER", "SERVICE_PERIOD", "AMOUNT_DUE"],
    "resourcing_kaiser": ["COMPANY_NAME", "BILLING_ID", "GROUP_ID", "INVOICE_NUMBER", "BILLING_MONTH", "TOTAL_AMOUNT_DUE"],
    "resourcing_uhc":    ["COMPANY_NAME", "INVOICE_NUMBER", "INVOICE_DATE", "COVERAGE_PERIOD", "CUSTOMER_NUMBER", "TOTAL_BALANCE_DUE"],
}

SUB_TYPE_LABELS = {
    "engage":            "Engage  (ADP TotalSource)",
    "velocity":          "Velocity  (Paychex PEO)",
    "prestige":          "Prestige  (Aetna)",
    "resourcing_kaiser": "Resourcing  (Kaiser Permanente)",
    "resourcing_uhc":    "Resourcing  (UnitedHealthcare)",
}

HEADER_COLOURS = {
    "engage":            "3C3489",
    "velocity":          "854F0B",
    "prestige":          "712B13",
    "resourcing_kaiser": "085041",
    "resourcing_uhc":    "0C447C",
}

# ══════════════════════════════════════════════════════════════════════════════
# KEYWORD CLASSIFIER
# ══════════════════════════════════════════════════════════════════════════════

KEYWORDS = {
    "engage":            ["TOTALSOURCE", "TOTALSOURCE BENEFITS INVOICE", "TOTALSOURCE® BENEFITS INVOICE", "NCT3-EPO"],
    "velocity":          ["PAYCHEX", "PEO BENEFITS ADMINISTRATION", "HUMAN RESOURCE SERVICES", "1175 JOHN ST"],
    "prestige":          ["AETNA", "CURRENT INFORCE CHARGES", "TRIAD NUMBER", "BILL PACKAGE"],
    "resourcing_kaiser": ["KAISER PERMANENTE", "KAISER FOUNDATION HEALTH PLAN", "BUSINESS.KP.ORG"],
    "resourcing_uhc":    ["UNITEDHEALTHCARE", "UHS PREMIUM BILLING", "UHCESERVICES.COM", "CONSOLIDATED CUSTOMER NO"],
}

# ══════════════════════════════════════════════════════════════════════════════
# LLM PROMPTS - output keys match EMPLOYEE_FIELDS exactly
# ══════════════════════════════════════════════════════════════════════════════

PROMPTS = {

    "engage": """
You are extracting data from an ADP TotalSource Benefits Invoice PDF.

Extract a SUMMARY and EMPLOYEES array.

SUMMARY: company_name, invoice_number, billing_date, due_date, reference_number, total_cost

EMP_LIST = [{"first_name":"","last_name":"","coverage":"","plan_name":"","coverage_option":"","current_premium":""}]
  employees - one row per plan line per employee, PLUS a subtotal row for each employee:
  first_name       : member first name
  last_name        : member last name
  coverage         : EXACT coverage tier/code (e.g. Employee, Family, EE+1, E, ES, ESC, EC, E1D, E2D, E3D, E4D, E5D, E6D, E7D, E8D, E9D, etc.)
  plan_name        : insurance category/type (e.g. Medical, Dental, Vision)
  coverage_option  : specific insurance product name (e.g. UnitedHealthcare Dental PPO 50, Choice Plus HDHP 1700)
  current_premium  : dollar amount for that plan line

Rules: 
1. INDIVIDUAL PLAN LINES: Extract every plan line for an employee with its individual cost.
2. SUBTOTAL ROW: After all plan lines for an employee, add ONE subtotal row:
   - coverage_option: (Set to exactly "TOTAL")
   - current_premium: the sub-total amount for that specific employee.
3. Each person must be represented by their individual plan rows, followed by their "TOTAL" subtotal row.
Use "" for missing values. Return ONLY valid JSON.

{{
  "summary": {{"company_name":"","invoice_number":"","billing_date":"","due_date":"","reference_number":"","total_cost":""}},
  "employees": [{{"first_name":"","last_name":"","coverage":"","plan_name":"","coverage_option":"","current_premium":""}}]
}}

PDF TEXT: {text}
""",

    "velocity": """
You are extracting data from a Paychex PEO Benefits Administration statement PDF.

Extract a SUMMARY and EMPLOYEES array.

SUMMARY: company_name, client_number, statement_date, statement_number, grand_total

EMPLOYEES - one row per plan line per employee, PLUS a subtotal row for each employee:
  first_name       : member first name
  last_name        : member last name
  coverage         : EXACT coverage tier/code (e.g. Employee, Family, EE+1, E, ES, ESC, EC, E1D, E2D, E3D, E4D, E5D, E6D, E7D, E8D, E9D, etc.)
  plan_name        : insurance category/type (e.g. Medical, Dental, Vision)
  coverage_option  : specific insurance product name (e.g. UnitedHealthcare Dental PPO 50, Choice Plus HDHP 1700)
  current_premium  : dollar amount for that plan line

Rules: 
1. INDIVIDUAL PLAN LINES: Extract every plan line for an employee with its individual cost.
2. SUBTOTAL ROW: After all plan lines for an employee, add ONE subtotal row:
   - coverage_option: (Set to exactly "TOTAL")
   - current_premium: the sub-total amount for that specific employee.
3. Each person must be represented by their individual plan rows, followed by their "TOTAL" subtotal row.
Use "" for missing values. Return ONLY valid JSON.

{{
  "summary": {{"company_name":"","client_number":"","statement_date":"","statement_number":"","grand_total":""}},
  "employees": [{{"first_name":"","last_name":"","coverage":"","plan_name":"","coverage_option":"","current_premium":""}}]
}}

PDF TEXT: {text}
""",

    "prestige": """
You are extracting data from an Aetna group health insurance bill PDF.

Extract a SUMMARY and EMPLOYEES array.

SUMMARY: company_name, bill_number, triad_number, account_number, service_period, amount_due

EMPLOYEES - one row per member (9 required fields only):
  data_row                  : sequential row number
  first_name                : member first name
  last_name                 : member last name
  gender                    : M or F
  relationship_to_employee  : EE / SP / CH
  dependent_of_employee_row : data_row of their employee (blank if EE)
  coverage                  : Extract EXACT coverage code from invoice (e.g. E, ES, ESC, EC, E1D, E2D, E3D, E4D, E5D, E6D, E7D, E8D, E9D, EE, FAM, etc.)
  cobra_participant          : Y or N
  plan_name                 : specific plan name enrolled

DO NOT include date_of_birth or home_zip_code - not required for Prestige.
Use "" for missing values. Return ONLY valid JSON.

{{
  "summary": {{"company_name":"","bill_number":"","triad_number":"","account_number":"","service_period":"","amount_due":""}},
  "employees": [{{"data_row":"","first_name":"","last_name":"","gender":"","relationship_to_employee":"","dependent_of_employee_row":"","coverage":"","cobra_participant":"","plan_name":""}}]
}}

PDF TEXT: {text}
""",

    "resourcing_kaiser": """
You are extracting data from a Kaiser Permanente group health bill PDF.

Extract a SUMMARY and EMPLOYEES array.

SUMMARY: company_name, billing_id, group_id, invoice_number, billing_month, total_amount_due

EMPLOYEES - one row per member (12 required fields only):
  data_row                  : sequential row number
  first_name                : member first name
  last_name                 : member last name
  gender                    : M or F
  date_of_birth             : date of birth (4-digit year format)
  home_zip_code             : member zip code
  relationship_to_employee  : EE / SP / CH
  dependent_of_employee_row : data_row of their employee (blank if EE)
  coverage                  : Extract EXACT coverage code from invoice (e.g. E, ES, ESC, EC, E1D, E2D, E3D, E4D, E5D, E6D, E7D, E8D, E9D, EE, FAM, etc.)
  cobra_participant          : Y or N
  current_plan_enrolled     : specific plan enrolled (CRITICAL: Plan names often span across MULTIPLE physical lines in the invoice. You MUST aggressively capture the entire multi-line block into one string. Do not stop at the first line! Enter for Employee row only, blank for dependent rows)
  monthly_total_premium     : total premium for employee's tier (enter for Employee row only, blank for dependent rows)

Use "" for missing values. Return ONLY valid JSON.

{{
  "summary": {{"company_name":"","billing_id":"","group_id":"","invoice_number":"","billing_month":"","total_amount_due":""}},
  "employees": [{{"data_row":"","first_name":"","last_name":"","gender":"","date_of_birth":"","home_zip_code":"","relationship_to_employee":"","dependent_of_employee_row":"","coverage":"","cobra_participant":"","current_plan_enrolled":"","monthly_total_premium":""}}]
}}

PDF TEXT: {text}
""",

    "resourcing_uhc": """
You are extracting data from a UnitedHealthcare (UHC) group insurance invoice PDF.

Extract a SUMMARY and EMPLOYEES array.

SUMMARY: company_name, invoice_number, invoice_date, coverage_period, customer_number, total_balance_due

EMPLOYEES - one row per member (12 required fields only):
  data_row                  : sequential row number
  first_name                : member first name
  last_name                 : member last name
  gender                    : M or F
  date_of_birth             : date of birth (4-digit year format)
  home_zip_code             : member zip code
  relationship_to_employee  : EE / SP / CH
  dependent_of_employee_row : data_row of their employee (blank if EE)
  coverage                  : Extract EXACT coverage code from invoice (e.g. E, ES, ESC, EC, E1D, E2D, E3D, E4D, E5D, E6D, E7D, E8D, E9D, EE, FAM, etc.)
  cobra_participant          : Y or N  (C = Cobra -> Y)
  current_plan_enrolled     : specific plan enrolled (CRITICAL: Plan names often print across MULTIPLE consecutive physical lines in the invoice. You MUST aggressively capture every single line of the plan description into one combined string. Do not stop at the first line! Enter for Employee row only, blank for dependent rows)
  monthly_total_premium     : total premium for employee's tier (enter for Employee row only, blank for dependent rows)

Use "" for missing values. Return ONLY valid JSON.

{{
  "summary": {{"company_name":"","invoice_number":"","invoice_date":"","coverage_period":"","customer_number":"","total_balance_due":""}},
  "employees": [{{"data_row":"","first_name":"","last_name":"","gender":"","date_of_birth":"","home_zip_code":"","relationship_to_employee":"","dependent_of_employee_row":"","coverage":"","cobra_participant":"","current_plan_enrolled":"","monthly_total_premium":""}}]
}}

PDF TEXT: {text}
""",
}

# ══════════════════════════════════════════════════════════════════════════════
# CORE FUNCTIONS
# ══════════════════════════════════════════════════════════════════════════════

def extract_text(pdf_path: Path, max_pages: int = 1000) -> str:
    """
    Robustly extracts text from a PDF by iterating through pages and using 
    multiple engines (pdfplumber, fitz, and OCR) to ensure full capture.
    Handles rotated and reversed text mapping issues.
    """
    text = ""
    print(f"[RPVE] Extracting text from {pdf_path.name}...")

    # Keywords we EXPECT to find in a valid RPVE document page
    VALID_KEYWORDS = ["TOTALSOURCE", "PAYCHEX", "AETNA", "KAISER", "UNITEDHEALTHCARE", "INVOICE", "BILLING", "PREMIUM", "AMOUNT DUE", "PAGE", "EMPLOYEE", "MEMBERS"]

    try:
        import pdfplumber
        import fitz
        import pytesseract
        from PIL import Image
        
        with pdfplumber.open(str(pdf_path)) as pdf:
            with fitz.open(pdf_path) as doc:
                pages_to_extract = min(max_pages, len(pdf.pages))
                for i in range(pages_to_extract):
                    page_text = ""
                    
                    # 1. Try pdfplumber (best for layout preservation)
                    try:
                        plumber_page = pdf.pages[i]
                        p_text = plumber_page.extract_text(layout=True) or ""
                        # Validate that it's not reversed or garbage by checking keywords
                        if len(p_text.strip()) > 100 and any(kw in p_text.upper() for kw in VALID_KEYWORDS):
                            page_text = p_text
                    except Exception as e:
                        print(f"  [PAGE {i+1}] pdfplumber error: {e}")

                    # 2. Fallback to fitz (PyMuPDF) if pdfplumber is empty, sparse, or fails keywords
                    if not page_text.strip():
                        try:
                            fitz_page = doc[i]
                            f_text = fitz_page.get_text() or ""
                            if len(f_text.strip()) > 50 and any(kw in f_text.upper() for kw in VALID_KEYWORDS):
                                page_text = f_text
                        except Exception as e:
                            print(f"  [PAGE {i+1}] fitz failed: {e}")

                    # 3. Last Resort: Per-page High-Accuracy OCR (handles scanned/rotated/reversed text)
                    if not page_text.strip():
                        print(f"  [PAGE {i+1}] No valid text layer detected (possibly scanned or reversed mapping). Running OCR...")
                        try:
                            # Render page at 2x zoom (144 dpi) for high accuracy OCR
                            pix = doc[i].get_pixmap(matrix=fitz.Matrix(2, 2))
                            img = Image.frombytes('RGB', [pix.width, pix.height], pix.samples)
                            # Use PSM 6 to assume a single uniform block of text
                            page_text = pytesseract.image_to_string(img, config='--psm 6')
                        except Exception as e:
                            print(f"  [PAGE {i+1}] OCR failed: {e}")

                    text += page_text + "\n"
    except Exception as e:
        print(f"[RPVE] Global extraction error: {e}")

    return text


def classify(text: str) -> str | None:
    t = text.upper()
    for sub_type in ["engage", "velocity", "prestige", "resourcing_kaiser", "resourcing_uhc"]:
        if any(kw in t for kw in KEYWORDS[sub_type]):
            return sub_type
    return None


def extract_with_llm(sub_type: str, text: str, ev_mode: bool = False) -> dict:
    # Increased input length limit to 400k to fully leverage the model's 128k token context window
    
    if ev_mode:
        # EV MODE IS ON: Strictly follow the 5-field requirement.
        # This applies regardless of the sub_type (Engage, Velocity, etc.)
        sum_fields = ", ".join(SUMMARY_FIELDS.get(sub_type, ["company_name", "total_amount_due"]))
        prompt_template = f"""
You are extracting data from a {sub_type.replace('_', ' ')} invoice.

Extract a SUMMARY and EMPLOYEES array.

SUMMARY: {sum_fields}

  employees: one row per individual plan line, PLUS a subtotal row for each employee:
  1. MULTI-PLAN BLOCKS: Employees often have 3 to 10 plans listed (Medical, Dental, Vision, Life, etc.).
  2. NAME ASSOCIATION: The employee name usually only appears on the FIRST plan line of their block. You MUST associate ALL subsequent plan lines with that same employee name.
  3. CATEGORY vs PRODUCT: 
     - plan_name: insurance category (e.g. Medical, Dental, Vision)
     - coverage_option: specific insurance product name
  4. ONE ROW PER PLAN: Each plan line MUST be its own object in the 'employees' list.
  5. SUBTOTAL ROW: After all plan lines for an employee, add ONE subtotal row:
     - coverage_option: (Set to exactly "TOTAL")
     - current_premium: the combined sub-total amount for that specific employee.
  6. Return ONLY valid JSON.

  FORMAT EXAMPLE:
  {{
    "summary": {{ "company_name": "...", "total_cost": "..." }},
    "employees": [
      {{"first_name": "John", "last_name": "Doe", "coverage": "EE", "plan_name": "Dental", "coverage_option": "UHC Dental PPO 50", "current_premium": "40.00"}},
      {{"first_name": "John", "last_name": "Doe", "coverage": "EE", "plan_name": "Vision", "coverage_option": "VSP Vision Plan", "current_premium": "10.00"}},
      {{"first_name": "John", "last_name": "Doe", "coverage": "", "plan_name": "", "coverage_option": "TOTAL", "current_premium": "50.00"}}
    ]
  }}

PDF TEXT: {{text}}
"""
    else:
        # EV MODE IS OFF: Strictly follow the 13-field requirement (Resourcing/Prestige).
        # Even if the document is Engage/Velocity, we force the 13-field structure.
        prompt_template = f"""
You are extracting data from a {sub_type.replace('_', ' ')} invoice.
EV MODE IS OFF: You must extract the following 14 fields for each person:
- Data Row (data_row)
- First Name (first_name)
- Last Name (last_name)
- Gender (gender - M / F)
- Date of Birth (date_of_birth)
- Home Zip Code (home_zip_code)
- Relationship to Employee (relationship_to_employee - EE / SP / CH)
- Dependent of Employee Row (dependent_of_employee_row)
- Coverage (coverage - ES / EC / FAM / NE / WP / RC / WO)
- Coverage Option (coverage_option - specific insurance product name e.g. Dental PPO 50, Choice Plus HDHP 1700)
- Cobra Participant (cobra_participant - Y / N)
- Current Plan Enrolled (current_plan_enrolled)
- Plan Name (plan_name)
- Monthly Total Premium (monthly_total_premium - only employee total, NOT dependents)

STRICT RULES:
1. One row per individual.
2. If the document has multiple plan lines for one person (like Engage), squash them into one row and use the total premium. You can comma-separate coverage_option if necessary.
3. Return valid JSON only.

{{
  "summary": {{"company_name": "", "total_amount_due": ""}},
  "employees": [{{"data_row": null, "first_name": null, "last_name": null, "gender": null, "date_of_birth": null, "home_zip_code": null, "relationship_to_employee": null, "dependent_of_employee_row": null, "coverage": null, "coverage_option": null, "cobra_participant": null, "current_plan_enrolled": null, "plan_name": null, "monthly_total_premium": null}}]
}}

PDF TEXT: {{text}}
"""

    lines = text.split('\n')
    chunks = []
    current_chunk = []
    current_len = 0
    # Chunk by ~40,000 chars to avoid hitting the 16k output tokens max limit and timeouts
    for line in lines:
        if current_len + len(line) > 40000 and current_chunk:
            chunks.append('\n'.join(current_chunk))
            current_chunk = []
            current_len = 0
        current_chunk.append(line)
        current_len += len(line) + 1
    if current_chunk:
        chunks.append('\n'.join(current_chunk))

    all_employees = []
    final_summary = {}

    print(f"[RPVE] LLM extraction split into {len(chunks)} chunks...")
    for i, chunk in enumerate(chunks):
        prompt = prompt_template.replace("{text}", chunk)
        try:
            response = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": "You are a precise insurance billing data extraction assistant. Return valid JSON only. No markdown. No extra text."},
                    {"role": "user", "content": prompt},
                ],
                temperature=0,
                response_format={"type": "json_object"},
            )
            raw = response.choices[0].message.content
            try:
                data = json.loads(raw)
            except json.JSONDecodeError:
                data = json.loads(re.sub(r"```json|```", "", raw).strip())
                
            if not final_summary and data.get("summary"):
                # Always grab summary from the first successful chunk
                final_summary = data.get("summary")
                
            emps = data.get("employees", [])
            all_employees.extend(emps)
            print(f"  [RPVE] Chunk {i+1}/{len(chunks)} processed -> found {len(emps)} employees")
        except Exception as e:
            print(f"  [RPVE] Chunk {i+1}/{len(chunks)} failed: {e}")

    return {
        "summary": final_summary,
        "employees": all_employees
    }


def build_excel(data: dict, sub_type: str, stem: str, active_employee_fields: list[str] | None = None) -> Path:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb        = Workbook()
    hex_col   = HEADER_COLOURS.get(sub_type, "1A1A2E")
    hdr_fill  = PatternFill("solid", fgColor=hex_col)
    hdr_font  = Font(color="FFFFFF", bold=True, size=11, name="Calibri")
    hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin      = Side(style="thin", color="DDDDDD")
    bdr       = Border(left=thin, right=thin, top=thin, bottom=thin)
    da        = Alignment(vertical="center")
    tf        = Font(bold=True, size=11, name="Calibri")
    tfill     = PatternFill("solid", fgColor="F0F0F0")

    summary   = data.get("summary", {})
    employees = data.get("employees", [])

    # ── Sheet 1: Employee Details ─────────────────────────────────────────────
    we = wb.active
    we.title = "Employee Details"
    we.sheet_view.showGridLines = False

    # Use passed fields or fallback to global mapping
    all_cols = active_employee_fields if active_employee_fields is not None else EMPLOYEE_FIELDS.get(sub_type, [])

    we.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(len(all_cols),1))
    t2 = we.cell(row=1, column=1, value=f"Employee Details - {len(employees)} records")
    t2.font      = Font(bold=True, size=13, color="FFFFFF", name="Calibri")
    t2.fill      = PatternFill("solid", fgColor=hex_col)
    t2.alignment = Alignment(horizontal="center", vertical="center")
    we.row_dimensions[1].height = 28

    for ci, col in enumerate(all_cols, 1):
        c = we.cell(row=2, column=ci, value=col.replace("_", " "))
        c.fill, c.font, c.alignment, c.border = hdr_fill, hdr_font, hdr_align, bdr
        we.column_dimensions[get_column_letter(ci)].width = 22
    we.row_dimensions[2].height = 22

    for ri, emp in enumerate(employees, 3):
        we.row_dimensions[ri].height = 18
        for ci, col in enumerate(all_cols, 1):
            c = we.cell(row=ri, column=ci, value=emp.get(col.lower(), ""))
            c.border, c.alignment = bdr, da
            c.font = Font(size=10, name="Calibri")
            if ri % 2 == 0:
                c.fill = PatternFill("solid", fgColor="F7F7F7")

    fin_cols = {"CURRENT_PREMIUM", "MONTHLY_TOTAL_PREMIUM", "GRAND_TOTAL", "TOTAL_COST"}
    fin_present = [c for c in all_cols if c in fin_cols]
    if fin_present and employees:
        tr = len(employees) + 3
        we.row_dimensions[tr].height = 20
        lc = we.cell(row=tr, column=1, value="TOTAL")
        lc.font, lc.fill, lc.border = tf, tfill, bdr
        for ci, col in enumerate(all_cols, 1):
            c = we.cell(row=tr, column=ci)
            c.fill, c.border = tfill, bdr
            if col in fin_cols:
                total = 0.0
                for emp in employees:
                    # Skip subtotal rows ("TOTAL" or empty plan_name) to avoid double counting
                    p_opt = emp.get("coverage_option", "")
                    pname = str(p_opt if p_opt is not None else "").strip().upper()
                    
                    if pname == "TOTAL" or (sub_type in ["engage", "velocity"] and (not pname or pname == "NONE")):
                        continue
                    
                    val_str = str(emp.get(col.lower(), "")).replace("$", "").replace(",", "")
                    try:
                        clean_val = re.sub(r'[^\d.-]', '', val_str)
                        if clean_val:
                            total += float(clean_val)
                    except:
                        pass
                c.value = f"${total:,.2f}"
                c.font  = tf

    # ── Sheet 2: Summary ──────────────────────────────────────────────
    if summary:
        ws = wb.create_sheet(title="Summary")
        ws.sheet_view.showGridLines = False
        sum_cols = SUMMARY_FIELDS.get(sub_type, list(summary.keys()))
        
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(len(sum_cols),1))
        tc = ws.cell(row=1, column=1, value=f"RPVE - {SUB_TYPE_LABELS.get(sub_type, sub_type.upper())}")
        tc.font  = Font(bold=True, size=13, color="FFFFFF", name="Calibri")
        tc.fill  = PatternFill("solid", fgColor=hex_col)
        tc.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 28

        for ci, col in enumerate(sum_cols, 1):
            c = ws.cell(row=2, column=ci, value=col.replace("_", " "))
            c.fill, c.font, c.alignment, c.border = hdr_fill, hdr_font, hdr_align, bdr
            ws.column_dimensions[get_column_letter(ci)].width = 24
        ws.row_dimensions[2].height = 22

        for ci, col in enumerate(sum_cols, 1):
            c = ws.cell(row=3, column=ci, value=summary.get(col.lower(), ""))
            c.border, c.alignment = bdr, da
            
        ws.freeze_panes = "A3"

    we.freeze_panes = "A3"
    wb.active = we

    xlsx_path = OUTPUT_DIR / f"{stem}_RPVE_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    wb.save(str(xlsx_path))
    print(f"[RPVE] Excel -> {xlsx_path.name}")
    return xlsx_path


def build_json_file(data: dict, sub_type: str, stem: str, active_employee_fields: list[str] | None = None) -> Path:
    summary   = data.get("summary", {})
    employees = data.get("employees", [])
    required  = active_employee_fields if active_employee_fields is not None else EMPLOYEE_FIELDS.get(sub_type, [])

    rows = []
    for emp in employees:
        row = {k.upper(): v for k, v in summary.items()}
        # Only include required fields - strip everything else
        for col in required:
            row[col] = emp.get(col.lower(), "")
        row["RPVE_SUB_TYPE"] = sub_type.upper()
        rows.append(row)

    json_path = OUTPUT_DIR / f"{stem}_RPVE_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    with open(str(json_path), "w", encoding="utf-8") as f:
        json.dump(rows, f, indent=2, ensure_ascii=False)
    print(f"[RPVE] JSON  -> {json_path.name}")
    return json_path


# ══════════════════════════════════════════════════════════════════════════════
# FASTAPI APP
# ══════════════════════════════════════════════════════════════════════════════

app = FastAPI(
    title="RPVE - Benefit Invoice Extractor",
    description="Resourcing · Prestige · Velocity · Engage",
    version="1.0.0",
)

app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

_cache: dict[str, str] = {}

FRONTEND_DIST_DIR = BASE_DIR / "frontend" / "dist"

if (FRONTEND_DIST_DIR / "assets").exists():
    app.mount("/assets", StaticFiles(directory=FRONTEND_DIST_DIR / "assets"), name="assets")

@app.get("/", response_class=HTMLResponse, include_in_schema=False)
async def serve_ui():
    ui = FRONTEND_DIST_DIR / "index.html"
    if ui.exists():
        return HTMLResponse(content=ui.read_text(encoding="utf-8"))
    
    ui_fallback = BASE_DIR / "RPVE_ui.html"
    if ui_fallback.exists():
        return HTMLResponse(content=ui_fallback.read_text(encoding="utf-8"))
        
    return HTMLResponse("<h2>RPVE running</h2><p>Build the frontend first.</p><a href='/docs'>Swagger -></a>")


@app.get("/api/health")
async def health():
    return {"status": "ok", "service": "RPVE", "sub_types": list(KEYWORDS.keys())}


@app.post("/api/extract")
async def extract(file: UploadFile = File(...), ev_mode: str = Form("false")):
    ev_bool = ev_mode.lower() == "true"
    print(f"\n[RPVE] Extraction Mode -> EV={ev_bool} (Raw: {ev_mode})")
    if not file.filename:
        raise HTTPException(400, "No filename provided")
    ext = Path(file.filename).suffix.lower()
    if ext not in [".pdf", ".csv", ".xlsx", ".xls"]:
        raise HTTPException(400, f"Supported formats: PDF, CSV, XLSX, XLS. Got: {ext}")

    safe = re.sub(r'[\\/:*?"<>|]', "_", file.filename)
    file_path = UPLOAD_DIR / safe
    with open(file_path, "wb") as f:
        shutil.copyfileobj(file.file, f)
    print(f"\n[RPVE] Received -> {safe}")

    try:
        from identification import universal_extract_text, ai_classify
        text = universal_extract_text(file_path)
        
        # Consistent Text Output: Save the extracted text for ALL file types
        txt_path = OUTPUT_DIR / f"{file_path.stem}.txt"
        with open(txt_path, "w", encoding="utf-8") as f:
            f.write(text)
        print(f"[RPVE] Saved structured text to {txt_path.name}")
        
    except Exception as read_err:
        print(f"[RPVE] Read error: {read_err}")
        raise HTTPException(422, f"Failed to extract text from {ext} file: {read_err}")

    if not text.strip():
        raise HTTPException(422, "No text extracted. File may be empty or an unreadable image.")

    sub_type = ai_classify(text, ev_mode=ev_bool)
    if sub_type is None:
        raise HTTPException(422, "AI Identification Failed: Document does not match any valid RPVE category.")

    print(f"[RPVE] AI Classified -> {sub_type}")

    try:
        data = extract_with_llm(sub_type, text, ev_mode=ev_bool)
    except Exception as e:
        raise HTTPException(500, f"LLM extraction failed: {str(e)}")

    emp_count = len(data.get("employees", []))
    print(f"[RPVE] Extracted  -> {emp_count} rows")

    stem      = Path(safe).stem
    
    try:
        # Determine the actual fields used for this extraction strictly by mode
        if ev_bool:
            active_fields = EV_ON_FIELDS
        else:
            active_fields = EV_OFF_FIELDS

        xlsx_path = build_excel(data, sub_type, stem, active_employee_fields=active_fields)
        json_path = build_json_file(data, sub_type, stem, active_employee_fields=active_fields)
    except Exception as build_err:
        import traceback
        print(f"[RPVE] Output building error:\n{traceback.format_exc()}")
        raise HTTPException(500, f"Failed to generate output files: {str(build_err)}")

    _cache[xlsx_path.name] = str(xlsx_path)
    _cache[json_path.name] = str(json_path)

    summary_dict = data.get("summary", {})
    total_val_str = "0"
    # Search all possible total keys in summary
    for tk in ["total_cost", "grand_total", "total_amount_due", "total_balance_due", "amount_due", "total_amount"]:
        val = summary_dict.get(tk) or summary_dict.get(tk.upper())
        if val:
            total_val_str = val
            break
            
    try:
        # Clean string: remove $, commas, etc.
        numeric_total = float(re.sub(r'[^0-9\.]', '', str(total_val_str)))
    except:
        numeric_total = 0.0

    return {
        "status":         "success",
        "type":           "INVOICE",
        "sub_type":       sub_type,
        "sub_type_label": SUB_TYPE_LABELS.get(sub_type, sub_type),
        "employee_count": emp_count,
        "fields_in_excel": EMPLOYEE_FIELDS.get(sub_type, []),
        "summary":        summary_dict,
        "excel_file":     xlsx_path.name,
        "json_file":      json_path.name,
        "output_file":    xlsx_path.name,
        "output_json":    json_path.name,
        "total_value":    numeric_total,
        "excel_url":      f"/api/download/{xlsx_path.name}",
        "json_url":       f"/api/download/{json_path.name}",
        "employees":      [
            {col: emp.get(col.lower(), "") for col in active_fields}
            for emp in data.get("employees", [])
        ],
    }


@app.get("/api/download/{filename}", include_in_schema=False)
async def download(filename: str):
    fp = Path(_cache.get(filename, OUTPUT_DIR / filename))
    if not fp.exists():
        raise HTTPException(404, f"File not found: {filename}")
    mt = ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          if filename.endswith(".xlsx") else "application/json")
    return FileResponse(path=fp, filename=filename, media_type=mt)


@app.get("/{filename}", include_in_schema=False)
async def serve_root_files(filename: str):
    if filename.startswith("api") or filename in ["docs", "openapi.json"]:
        raise HTTPException(404)
    file_path = FRONTEND_DIST_DIR / filename
    if file_path.exists() and file_path.is_file():
        return FileResponse(file_path)
    raise HTTPException(404)


if __name__ == "__main__":
    import uvicorn, sys
    port = 8009
    if "--port" in sys.argv:
        try: port = int(sys.argv[sys.argv.index("--port") + 1])
        except: pass

    print("\n" + "="*50)
    print("  RPVE - Benefit Invoice Extractor")
    print("="*50)
    print(f"  UI      ->  http://localhost:{port}")
    print(f"  Swagger ->  http://localhost:{port}/docs")
    print("="*50 + "\n")
    uvicorn.run("RPVE_standalone:app", host="0.0.0.0", port=port, reload=True)
