import sys
from pathlib import Path

sys.path.append(str(Path(__file__).parent))

from identification import universal_extract_text, ai_classify
from RPVE_standalone import extract_with_llm, deduplicate_employees

def test_file(file_path_str):
    path = Path(file_path_str)
    print(f"--- Extracting {path.name} ---")
    text = universal_extract_text(path)
    
    sub_type = ai_classify(text) or "standard"
    print(f"Sub-type determined as: {sub_type}")
    
    data = extract_with_llm(sub_type, text)
    emps = data.get("employees", [])
    
    print(f"Raw employees extracted: {len(emps)}")
    
    valid_emps = []
    for e in emps:
        fname = str(e.get("first_name") or "").strip()
        lname = str(e.get("last_name") or "").strip()
        fulln = str(e.get("full_name") or "").strip()
        if (fname and lname) or fulln:
            valid_emps.append(e)
            
    dedup = deduplicate_employees(valid_emps)
    print(f"Valid and deduplicated employees: {len(dedup)}")
    
    if len(dedup) > 0:
        print("First 3:")
        for e in dedup[:3]:
            print(f"  {e.get('first_name')} {e.get('last_name')} | Plan: {e.get('plan_name')} | Type: {e.get('plan_type')} | Cvg: {e.get('coverage')} | Premium: {e.get('current_premium')} | Adj: {e.get('adjustment_amount')}")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        test_file(sys.argv[1])
    else:
        test_file(r"c:\Users\Intern\rvpe\MANIFEST MEDEX.xlsx")
        print("\n")
        test_file(r"c:\Users\Intern\rvpe\Bluecross Blue shield of north carolina.pdf")
        print("\n")
        test_file(r"c:\Users\Intern\rvpe\(Advantex Solutions Inc  )ADP Med Invoice 3-13-2026 1.pdf")
        print("\n")
        test_file(r"c:\Users\Intern\rvpe\CENTER FOR HUMAN DEVELOPMENT & FAMILY SERVICES CENTER .invoice.pdf")
