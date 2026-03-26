import os
import json
import traceback
from pathlib import Path
from RPVE_standalone import extract_with_llm

text_path = Path(r"c:\Users\INT002\RPVE\rpve_outputs\Benefits Billing 3-19-2026.txt")
text = text_path.read_text(encoding="utf-8")
print(f"Total Text Length: {len(text)} chars")

try:
    data = extract_with_llm("prestige", text, ev_mode=True)
    print("Success! Extracted rows:", len(data.get("employees", [])))
except Exception as e:
    print("Extraction Failed!")
    traceback.print_exc()
