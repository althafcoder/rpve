import asyncio
import os
import sys
from pathlib import Path

# Add project root to sys.path
sys.path.append(str(Path(__file__).parent.parent))
import RPVE_standalone

async def test_extraction():
    pdf_path = Path(r"c:\Users\Intern\rvpe\rpve\rpve_uploads\e8b8ae6b_20260515_175135_1pdf.pdf")
    if not pdf_path.exists():
        print(f"File not found: {pdf_path}")
        return

    print(f"Testing extraction for {pdf_path.name}...")
    try:
        result = await RPVE_standalone.process_invoice_data(pdf_path, pdf_path.name)
        print(f"Extraction successful! Found {result['employee_count']} records.")
        
        # Print first few records
        for i, emp in enumerate(result['employees'][:5]):
            print(f"  [{i+1}] {emp.get('full_name')} - {emp.get('plan_name')} - {emp.get('current_premium')}")
            
    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"Extraction failed: {e}")

if __name__ == "__main__":
    asyncio.run(test_extraction())
