import os
import csv
import json
from pathlib import Path
import openpyxl
from openai import OpenAI
from dotenv import load_dotenv

load_dotenv()
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

# Import the existing robust PDF extractor
from RPVE_standalone import extract_text as pdf_extract_text, OUTPUT_DIR

def universal_extract_text(file_path: Path) -> str:
    """Extracts raw text from PDF, CSV, or XLSX/XLS files for AI classification."""
    ext = file_path.suffix.lower()
    text = ""
    
    if ext == ".pdf":
        from RPVE_standalone import extract_text as pdf_extract_text
        text = pdf_extract_text(file_path)
    
    elif ext == ".csv":
        print(f"[IDENTIFY] Extracting text from CSV: {file_path.name}")
        try:
            with open(file_path, "r", encoding="utf-8-sig") as f:
                reader = csv.reader(f)
                for i, row in enumerate(reader):
                    if i > 1000: # Increased limit
                        break
                    # PRESERVE EMPTY COLUMNS: Map None to "" and join everything
                    text += " | ".join(str(cell).strip() if cell is not None else "" for cell in row) + "\n"
        except Exception as e:
            print(f"[IDENTIFY] CSV Error: {e}")
            
    elif ext in [".xlsx", ".xls"]:
        print(f"[IDENTIFY] Extracting text from Excel ({ext}): {file_path.name}")
        try:
            if ext == ".xlsx":
                wb = openpyxl.load_workbook(file_path, data_only=True)
                ws = wb.active
                for i, row in enumerate(ws.iter_rows(values_only=True)):
                    if i > 1000: break
                    # PRESERVE EMPTY COLUMNS: Map None to "" and join everything
                    text += " | ".join(str(cell).strip() if cell is not None else "" for cell in row) + "\n"
            else: # .xls
                import xlrd
                wb = xlrd.open_workbook(file_path)
                ws = wb.sheet_by_index(0)
                for i in range(min(ws.nrows, 1000)):
                    row = ws.row_values(i)
                    # PRESERVE EMPTY COLUMNS: Map None to "" and join everything
                    text += " | ".join(str(cell).strip() if cell is not None else "" for cell in row) + "\n"
        except Exception as e:
            print(f"[IDENTIFY] Excel Error: {e}")
            
    else:
        raise ValueError(f"Unsupported file format for extraction: {ext}")
        
    return text


def ai_classify(text: str, ev_mode: bool = False) -> str:
    """
    Uses fast-pass strict overrides for guaranteed accuracy on known carriers, 
    and falls back to an LLM to smartly categorize unknown/new documents.
    Returns one of: engage, velocity, prestige, resourcing_kaiser, resourcing_uhc.
    Returns None if absolutely no match.
    """
    from RPVE_standalone import KEYWORDS
    t_upper = text.upper()
    
    allowed_categories = ["engage", "velocity", "prestige", "resourcing_kaiser", "resourcing_uhc"]
    
    # 1. Fast-pass strict keyword overrides for guaranteed 100% accuracy
    if "TRINET" in t_upper or "PAYCHEX" in t_upper:
        print("[IDENTIFY] Fast-Pass Override: Identified as Velocity")
        return "velocity"
    if "TOTALSOURCE" in t_upper or "NCT3-EPO" in t_upper:
        print("[IDENTIFY] Fast-Pass Override: Identified as Engage")
        return "engage"
    if "PRESTIGE" in t_upper or "TRIAD NUMBER" in t_upper:
        print("[IDENTIFY] Fast-Pass Override: Identified as Prestige")
        return "prestige"
    if "KAISER PERMANENTE" in t_upper or "KAISER FOUNDATION" in t_upper:
        print("[IDENTIFY] Fast-Pass Override: Identified as Resourcing (Kaiser)")
        return "resourcing_kaiser"
    if "UNITEDHEALTHCARE" in t_upper and "UHS PREMIUM BILLING" in t_upper:
        print("[IDENTIFY] Fast-Pass Override: Identified as Resourcing (UHC)")
        return "resourcing_uhc"

    print("[IDENTIFY] Running AI-powered Classification fallback...")

    
    # We only need the first ~5000 characters to figure out what kind of invoice it is
    preview_text = text[:5000]
    
    prompt = f"""
You are an expert health insurance invoice classifier.
Analyze the following document text and determine which of our 5 internal categories it strictly maps to.

CATEGORIES AND CARRIER CLUES:
1. "engage" -> Usually ADP TotalSource, Engage PEO, NCT3-EPO, etc.
2. "velocity" -> Usually Paychex PEO, Human Resource Services, Trinet, Cigna, BlueCross BlueShield, Providence mapped under Velocity.
3. "prestige" -> Usually Aetna, Prestige, Triad Number, Bill Package.
4. "resourcing_kaiser" -> Kaiser Permanente, Kaiser Foundation.
5. "resourcing_uhc" -> UnitedHealthcare, UHS Premium Billing.

DOCUMENT PREVIEW:
{preview_text}

Respond IMMEDIATELY with a strictly formatted JSON object containing exactly one key "category" with one of the exact string values: "engage", "velocity", "prestige", "resourcing_kaiser", "resourcing_uhc" or `null` if it is completely unidentifiable. No markdown blocks, just pure JSON.

Example: {{"category": "velocity"}}
"""

    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            temperature=0.0,
            messages=[{"role": "user", "content": prompt}]
        )
        content = response.choices[0].message.content.strip()
        print(f"[IDENTIFY] AI Raw Output: {content}")
        
        if content.startswith("```json"):
            content = content[7:-3].strip()
        elif content.startswith("```"):
            content = content[3:-3].strip()
            
        result = json.loads(content)
        category = result.get("category")
        
        if category in ["engage", "velocity", "prestige", "resourcing_kaiser", "resourcing_uhc"]:
            return category
        else:
            print(f"[IDENTIFY] Invalid category returned: {category}")
            
    except Exception as e:
        print(f"[IDENTIFY] AI Classification failed: {e}")
        
    return None
