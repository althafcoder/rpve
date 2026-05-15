import sys
import os
from pathlib import Path

# Add the parent directory to sys.path
sys.path.append(str(Path(__file__).parent.parent))
from schema_ocr import SchemaOCRExtractor

pdf_path = Path(r"c:\Users\Intern\rvpe\rpve\rpve_uploads\e8b8ae6b_20260515_175135_1pdf.pdf")
if not pdf_path.exists():
    print(f"File not found: {pdf_path}")
    sys.exit(1)

extractor = SchemaOCRExtractor(pdf_path)
text = extractor.extract_layout_text(save_debug_output=True)

print("--- START OF TEXT ---")
print(text)
print("--- END OF TEXT ---")
