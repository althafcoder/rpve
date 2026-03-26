import os
import json
from pathlib import Path
from dotenv import load_dotenv

# Mocking FastAPI bits if needed, but let's just test the functions
from RPVE_standalone import extract_with_llm, build_excel, build_json_file, EMPLOYEE_FIELDS, EV_ON_FIELDS, EV_OFF_FIELDS
from identification import universal_extract_text, ai_classify

def debug():
    load_dotenv()
    file_path = Path(r"c:\Users\INT002\RPVE\rpve_uploads\Spur - Census.xlsx")
    stem = file_path.stem
    ev_mode = False # Default
    
    print(f"--- Debugging Extraction for {file_path.name} ---")
    
    try:
        print("1. Extracting text...")
        text = universal_extract_text(file_path)
        print(f"   Text length: {len(text)}")
        
        print("2. Classifying...")
        sub_type = ai_classify(text, ev_mode=ev_mode)
        print(f"   Sub-type: {sub_type}")
        
        if sub_type is None:
            print("   Classification failed!")
            return

        print("3. LLM Extraction...")
        data = extract_with_llm(sub_type, text, ev_mode=ev_mode)
        print("   Extraction successful.")
        
        print("4. Building outputs...")
        # Determine active fields
        active_fields = EV_ON_FIELDS if ev_mode else EV_OFF_FIELDS
        
        # Temporarily override EMPLOYEE_FIELDS
        original_fields = EMPLOYEE_FIELDS.copy()
        EMPLOYEE_FIELDS[sub_type] = active_fields
        
        try:
            xlsx_path = build_excel(data, sub_type, stem)
            print(f"   Excel built: {xlsx_path.name}")
            
            json_path = build_json_file(data, sub_type, stem)
            print(f"   JSON built: {json_path.name}")
        finally:
            # Restore
            EMPLOYEE_FIELDS.clear()
            EMPLOYEE_FIELDS.update(original_fields)
            
        print("--- DEBUG COMPLETE: SUCCESS ---")
        
    except Exception as e:
        print(f"--- DEBUG FAILED ---")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    debug()
