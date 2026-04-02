import requests
import json
import os

def test_api():
    url = "http://localhost:8011/api/extract"
    files_to_test = [
        r"c:\Users\Intern\rvpe\MANIFEST MEDEX.xlsx",
        r"c:\Users\Intern\rvpe\Bluecross Blue shield of north carolina.pdf",
        r"c:\Users\Intern\rvpe\(Advantex Solutions Inc  )ADP Med Invoice 3-13-2026 1.pdf",
        r"c:\Users\Intern\rvpe\CENTER FOR HUMAN DEVELOPMENT & FAMILY SERVICES CENTER .invoice.pdf",
        r"c:\Users\Intern\rvpe\G4 Geomatic Resources LLC invoice_ 1.pdf"
    ]
    
    for fpath in files_to_test:
        fname = os.path.basename(fpath)
        print(f"\nTesting {fname} ...")
        
        with open(fpath, "rb") as f:
            files = {"file": (fname, f)}
            try:
                response = requests.post(url, files=files, timeout=600)
                data = response.json()
                print(f"Status Code: {response.status_code}")
                if response.status_code == 200:
                    print(f"Employee Count: {data.get('employee_count')}")
                    emps = data.get("employees", [])
                    
                    if emps:
                        print("Sample Employee 1:")
                        print(json.dumps(emps[0], indent=2))
                else:
                    print(f"Error: {data}")
            except Exception as e:
                print(f"Request failed: {e}")

if __name__ == "__main__":
    test_api()
