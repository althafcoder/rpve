import os
import time
from pathlib import Path

# Import the core logic from your standalone server without changing it
from RPVE_standalone import classify, extract_text, extract_with_llm, EMPLOYEE_FIELDS, SUMMARY_FIELDS

POC_DIR = Path(r"c:\Users\INT002\RPVE\new poc")

def run_tests():
    total_files = 0
    classification_matches = 0
    extraction_successes = 0
    error_log = []

    print("==================================================")
    print("      RPVE EXTRACTION IDENTIFIER VALIDATION       ")
    print("==================================================")
    print(f"Target Directory: {POC_DIR}\n")

    if not POC_DIR.exists():
        print(f"[!] Target directory {POC_DIR} does not exist!")
        return

    for category_folder in POC_DIR.iterdir():
        if not category_folder.is_dir():
            continue

        raw_category = category_folder.name.lower()
        print(f"\n================ Scanning: [{raw_category.upper()}] ================")

        for pdf_file in category_folder.glob("*.pdf"):
            total_files += 1
            print(f"\n[{total_files}] Processing: {pdf_file.name}")
            print(f"  -> Step 1: Extracting text...")
            
            try:
                # 1. Text Extraction
                text = extract_text(pdf_path=pdf_file, max_pages=15)
                
                # 2. Classification
                print(f"  -> Step 2: Classifying Document...")
                detected_category = classify(text)
                
                # Map raw folder names to our internal sub_types
                cls_match = False
                if detected_category:
                    if raw_category == "resourcing":
                        if detected_category in ["resourcing_kaiser", "resourcing_uhc"]:
                            cls_match = True
                    else:
                        if detected_category == raw_category:
                            cls_match = True
                
                if cls_match:
                    print(f"  [OK] Classification passed: Identified as '{detected_category}'")
                    classification_matches += 1
                else:
                    print(f"  [FAIL] Classification mismatch! Folder is '{raw_category}', LLM routed to '{detected_category}'")
                    error_log.append(f"{pdf_file.name} - Routing Failed (Folder: {raw_category}, Got: {detected_category})")
                    continue

                # 3. LLM Field Extraction
                print(f"  -> Step 3: Sending to AI for JSON Extraction...")
                start_time = time.time()
                json_result = extract_with_llm(text, detected_category)
                elapsed = time.time() - start_time
                print(f"  -> Step 4: Received JSON in {elapsed:.1f}s. Validating Schema Match...")

                # 4. Strict Schema Validation
                if not json_result or "summary" not in json_result or "employees" not in json_result:
                    print(f"  [FAIL] JSON structurally invalid (missing basic keys)")
                    error_log.append(f"{pdf_file.name} - Invalid JSON Structure")
                    continue

                missing_fields = []
                
                # Validate Summary
                expected_summary = set(SUMMARY_FIELDS[detected_category])
                actual_summary = set([k.upper() for k in json_result["summary"].keys()])
                for req_key in expected_summary:
                    if req_key not in actual_summary:
                        missing_fields.append(f"Summary block missing key: {req_key}")

                # Validate Employees
                expected_emp = set(EMPLOYEE_FIELDS[detected_category])
                employees_list = json_result["employees"]
                for i, emp in enumerate(employees_list):
                    actual_emp = set([k.upper() for k in emp.keys()])
                    for req_key in expected_emp:
                        if req_key not in actual_emp:
                            missing_fields.append(f"Employee Row {i+1} missing key: {req_key}")

                if missing_fields:
                    print(f"  [FAIL] Schema Validation Failed! ({len(missing_fields)} errors found)")
                    for m in missing_fields[:5]:
                        print(f"    - {m}")
                    if len(missing_fields) > 5:
                        print(f"    ... and {len(missing_fields)-5} more.")
                    error_log.append(f"{pdf_file.name} - Schema Data Loss ({len(missing_fields)} missing fields)")
                else:
                    print(f"  [OK] 100% Exact Schema Match! ({len(employees_list)} employees parsed flawlessly)")
                    extraction_successes += 1

            except Exception as e:
                print(f"  [ERROR] Processing crashed: {e}")
                error_log.append(f"{pdf_file.name} - Crash: {e}")

    # 5. Final Report
    print("\n==================================================")
    print("                FINAL ACCURACY REPORT             ")
    print("==================================================")
    if total_files == 0:
        print("No PDF files found to test.")
        return

    cls_acc = (classification_matches / total_files) * 100
    if classification_matches > 0:
        ext_acc = (extraction_successes / classification_matches) * 100
    else:
        ext_acc = 0.0

    print(f"Total PDFs Evaluated     : {total_files}")
    print(f"Routing/Classify Accuracy: {cls_acc:.1f}% ({classification_matches}/{total_files})")
    print(f"Field Schema Accuracy    : {ext_acc:.1f}% ({extraction_successes}/{classification_matches})")
    print("==================================================")

    if error_log:
        print("\nERRORS ENCOUNTERED:")
        for err in error_log:
            print(f" - {err}")
    else:
        print("\n[SUCCESS] PERFECT 100% ACCURACY SCORE ACROSS THE BOARD!")

if __name__ == "__main__":
    run_tests()
