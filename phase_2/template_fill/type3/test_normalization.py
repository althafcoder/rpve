import os
import sys
import json
from pathlib import Path

# Add current dir to path
sys.path.append(str(Path(__file__).parent.parent.absolute()))

from type3 import census_normalizer

def test_normalization():
    path = r"C:\Users\Intern\rvpe\rpve\rpve_uploads\97ea9ab8_20260512_181919_PEO Spectrum RFP Employee Census Template - B2R.xlsx"
    
    if not os.path.exists(path):
        print(f"Error: File not found {path}")
        return

    print(f"Normalizing: {Path(path).name}")
    try:
        results = census_normalizer.normalize_census_to_list(path)
        print(f"Success! Normalized {len(results)} employees.")
        
        # Check specific employee: Mark Van Wonterghem
        # The key is 'first last' lowercase
        key = "mark van wonterghem"
        if key in results:
            emp = results[key]
            print(f"\nEmployee: {emp['first']} {emp['last']}")
            print(f"  Gender: {emp['gender']}")
            print(f"  DOB: {emp['dob']}")
            print(f"  Zip: {emp['zip']}")
            print(f"  Coverage: {emp['coverage']}")
            print(f"  Dependents: {len(emp['dependents'])}")
            for d in emp['dependents']:
                print(f"    - {d['first']} {d['last']} ({d['relation']}) DOB={d['dob']}")
        else:
            print(f"\nError: Could not find '{key}' in results.")
            print(f"Available keys: {list(results.keys())[:10]}...")

    except Exception as e:
        print(f"Error during normalization: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_normalization()
