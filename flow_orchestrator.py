import sys
import subprocess
import os
from pathlib import Path
from datetime import datetime
import pandas as pd
import asyncio

# Set up paths
CURRENT_DIR = Path(__file__).parent.absolute()
sys.path.append(str(CURRENT_DIR))

try:
    import RPVE_standalone
    from RPVE_standalone import process_invoice_data, OUTPUT_DIR
except ImportError as e:
    print(f"Error: Could not import RPVE_standalone.py. {e}")
    print(f"Current directory: {os.getcwd()}")
    print(f"Script directory: {CURRENT_DIR}")
    sys.exit(1)

def ensure_xlsx(file_path: str) -> str:
    """If file is .xls, convert to .xlsx using pandas and return new path."""
    if not file_path or not file_path.lower().endswith(".xls"):
        return file_path
        
    print(f"    -> Converting old .xls format to .xlsx: {Path(file_path).name}")
    try:
        new_path = str(Path(file_path).with_suffix(".xlsx"))
        import warnings
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            xl = pd.ExcelFile(file_path)
            with pd.ExcelWriter(new_path, engine='openpyxl') as writer:
                for sheet_name in xl.sheet_names:
                    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
                    df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
        return new_path
    except Exception as e:
        print(f"    -> Warning: Failed to convert .xls: {e}")
        return file_path

def classify_excel_template(excel_path: Path) -> str:
    """Scans the first 25 rows of an Excel file to fingerprint its Type."""
    try:
        # Suppress warnings for xls files
        import warnings
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            df = pd.read_excel(excel_path, nrows=25, header=None)
        all_text = " ".join([str(val).lower() for val in df.values.flatten() if pd.notna(val)])
        
        # Type 1 Detection (Engage/Kaiser)
        if "ee row" in all_text or "relation-ship to employee" in all_text or "kaiser networks" in all_text:
            return "type1"
            
        # Type 3 Detection (RAPT Blue Headers)
        elif "data row" in all_text or "cobra participant" in all_text or "discrepancies" in all_text:
            return "type3"
            
        # Type 2 (Basic Titan Intake / Generic Census)
        # Type 2 (`fill_template.py`) is designed as a fully dynamic universal filler.
        elif "first name" in all_text or "last name" in all_text or "name" in all_text:
            return "type2"
            
        # Default fallback to type2 because it uses dynamic column header mapping
        # and works for almost any standard census layout.
        return "type2"
    except Exception as e:
        print(f"Error reading Excel: {e}")
        return "type2"  # Fallback to dynamic filler instead of crashing

async def run_flow(pdf_path: str, template_path: str, ref_census_path: str = None):
    print(f"--- Starting End-to-End RPVE Flow ---")
    
    # ── PHASE 1: PDF Extraction ───────────────────────────────────────────────
    print(f"\n[1] Running Phase 1 (Extraction) on: {pdf_path}")
    try:
        p_path = Path(pdf_path)
        result_data = await process_invoice_data(p_path, p_path.name)
        phase1_output_excel = result_data['excel_path']
        print(f"    -> Phase 1 Success! Extracted: {phase1_output_excel}")
        
        output_dir = Path(phase1_output_excel).parent
        timestamp  = datetime.now().strftime('%Y%m%d_%H%M%S')

        phase2_report_name    = f"FINAL_AUDIT_REPORT_{timestamp}.xlsx"
        phase2_report_path    = output_dir / phase2_report_name

        validated_report_name = f"VALIDATED_AUDIT_REPORT_{timestamp}.xlsx"
        validated_report_path = output_dir / validated_report_name
        
    except Exception as e:
        print(f"    -> Phase 1 Failed: {e}")
        import traceback
        traceback.print_exc()
        traceback.print_exc()
        traceback.print_exc()
        return
    
    # ── PHASE 1.5: Pre-process Templates (.xls -> .xlsx) ──────────────────────
    template_path   = ensure_xlsx(template_path)
    ref_census_path = ensure_xlsx(ref_census_path)

    # ── PHASE 2: Template Fill ────────────────────────────────────────────────
    print(f"\n[2] Analyzing Template: {template_path}")
    
    template_type = classify_excel_template(Path(template_path))
    
    if ref_census_path:
        ref_type = classify_excel_template(Path(ref_census_path))
        
        # ── ROBUST SWAP INTELLIGENCE ────────────────────────────────────────
        t_name = Path(template_path).name.lower()
        r_name = Path(ref_census_path).name.lower()
        
        print(f"    -> Swap Check: Template={t_name}, Ref={r_name}")
        print(f"    -> Type Check: Template={template_type}, Ref={ref_type}")

        swapped = False
        # 1. Name-based hint: If the Reference slot contains "RAPT" and Template doesn't, swap.
        # BUT: Do not swap if the template already looks like a valid specific template (Type 1 or 3).
        if template_type not in ["type1", "type3"] and (("rapt" in r_name and "rapt" not in t_name) or ("template" in r_name and "template" not in t_name)):
            print(f"    -> Name-based Swap: User put RAPT/Template file in Reference slot. Swapping roles...")
            template_path, ref_census_path = ref_census_path, template_path
            # Re-classify after swap
            template_type = classify_excel_template(Path(template_path))
            ref_type = classify_excel_template(Path(ref_census_path))
            swapped = True
            print(f"    -> Post-Swap Types: Template={template_type}, Ref={ref_type}")

        # 2. Volume-based check: The 'Reference' file should be the source of truth (more data).
        if not swapped:
            try:
                import warnings
                with warnings.catch_warnings():
                    warnings.simplefilter("ignore")
                    t_df = pd.read_excel(template_path, nrows=100)
                    r_df = pd.read_excel(ref_census_path, nrows=100)
                    
                    # Count non-nulls in first 3 cols to ignore empty formatted templates
                    t_val_count = t_df.iloc[:, :3].count().sum()
                    r_val_count = r_df.iloc[:, :3].count().sum()
                    
                    print(f"    -> Content Check: Template={t_val_count} values, Ref={r_val_count} values")
                    
                    if t_val_count > (r_val_count + 10):
                        print(f"    -> Volume-based Swap Triggered!")
                        template_path, ref_census_path = ref_census_path, template_path
                        template_type = classify_excel_template(Path(template_path))
                        ref_type = classify_excel_template(Path(ref_census_path))
                        swapped = True
            except Exception as e:
                print(f"    -> Volume check failed: {e}")

        # 3. Format-based fallback
        if not swapped and ref_type == "type3" and template_type != "type3":
            print(f"    -> Format-based Swap: User put RAPT file in Reference slot. Swapping roles...")
            template_path, ref_census_path = ref_census_path, template_path
            template_type = "type3"
        elif ref_census_path:
            template_type = "type3" 
            
        print(f"    -> Final Decision: Template Type={template_type.upper()}")
    else:
        print(f"    -> Detected Template Type: {template_type.upper()}")
    
    print(f"\n[3] Executing Phase 2 (Template Fill)...")
    phase2_base = Path(__file__).parent / "phase_2" / "template_fill"
    
    if template_type == "type1":
        script = phase2_base / "type1" / "fill_template_v2.py"
        cmd = [sys.executable, str(script), phase1_output_excel, template_path, str(phase2_report_path)]
        
    elif template_type == "type2":
        script = phase2_base / "type2" / "fill_template.py"
        cmd = [sys.executable, str(script), phase1_output_excel, template_path, str(phase2_report_path)]
        
    elif template_type == "type3":
        if not ref_census_path:
            default_ref = phase2_base / "type3" / "reference_census" / "TEPCensus.xlsx"
            if default_ref.exists():
                ref_census_path = str(default_ref)
            else:
                print("Error: Type 3 requires a Reference Census file!")
                return
        script = phase2_base / "type3" / "fill_template.py"
        cmd = [sys.executable, str(script), phase1_output_excel, ref_census_path, template_path, str(phase2_report_path)]
        
    else:
        print("Error: Could not identify the Excel template type.")
        return

    print(f"    Running command: {' '.join(cmd)}")
    result = subprocess.run(cmd, capture_output=True, text=True)
    
    if result.returncode != 0:
        print("\n--- Flow Failed in Phase 2 ---")
        print(result.stderr)
        raise Exception(f"Phase 2 failed: {result.stderr}")

    print("    -> Phase 2 (Census Fill) Success!")
    if result.stdout:
        print(result.stdout)

    # ── PHASE 3: Data Validation (Name Normalisation + Fuzzy Resolve) ─────────
    print(f"\n[4] Executing Phase 3 (Data Validation — Name Normalisation)...")
    validation_script = phase2_base / "data_validation.py"

    if not validation_script.exists():
        print(f"    [WARN] data_validation.py not found at {validation_script}. Skipping Phase 3.")
        validated_report_path = phase2_report_path
        validated_report_name = phase2_report_name
    else:
        val_cmd = [
            sys.executable, str(validation_script),
            str(phase2_report_path),      # filled census (Phase 2 output)
            phase1_output_excel,          # invoice extraction (Phase 1 output)
            str(validated_report_path),   # validated output (Phase 3 final)
            "--threshold", "85",          # fuzzy match confidence threshold
            "--template-type", template_type,  # CH/SP skip only for type1 (Engage)
        ]
        print(f"    Running command: {' '.join(val_cmd)}")
        val_result = subprocess.run(val_cmd, capture_output=True, text=True)

        if val_result.returncode == 0:
            print("    -> Phase 3 (Data Validation) Success!")
            if val_result.stdout:
                print(val_result.stdout)
        else:
            # Phase 3 failure is non-fatal — fall back to Phase 2 output
            print(f"    [WARN] Phase 3 failed (non-fatal). Using Phase 2 output as final.")
            print(val_result.stderr)
            validated_report_path = phase2_report_path
            validated_report_name = phase2_report_name

    # ── PHASE 4: LLM Resolution (Fallback) ──────────────────────────────────
    print(f"\n[5] Executing Phase 4 (LLM Resolution)...")
    llm_script = phase2_base / "llm_resolution.py"
    audit_json_path = validated_report_path.with_suffix('.audit.json')
    llm_report_path = validated_report_path.with_name(validated_report_path.name.replace("VALIDATED_", "LLM_RESOLVED_"))
    llm_report_name = llm_report_path.name

    if not llm_script.exists() or not audit_json_path.exists():
        print(f"    [WARN] LLM resolution requirements not met. Skipping Phase 4.")
        final_report_path = validated_report_path
        final_report_name = validated_report_name
    else:
        llm_cmd = [
            sys.executable, str(llm_script),
            str(validated_report_path),
            str(audit_json_path),
            "--output", str(llm_report_path),
            "--template-type", template_type,   # pass type so LLM knows CH/SP rule applies only for type3
        ]
        print(f"    Running command: {' '.join(llm_cmd)}")
        llm_result = subprocess.run(llm_cmd, capture_output=True, text=True)
        
        if llm_result.returncode == 0 and llm_report_path.exists():
            print("    -> Phase 4 (LLM Resolution) Success!")
            if llm_result.stdout:
                print(llm_result.stdout)
            if llm_result.stderr:
                print(llm_result.stderr)
            final_report_path = llm_report_path
            final_report_name = llm_report_name
        else:
            print(f"    [WARN] Phase 4 failed or skipped. Using Phase 3 output as final.")
            if llm_result.stderr:
                print(llm_result.stderr)
            final_report_path = validated_report_path
            final_report_name = validated_report_name

    # ── Register all outputs in backend download cache ───────────────────────
    if 'RPVE_standalone' in sys.modules:
        _standalone = sys.modules['RPVE_standalone']
        if hasattr(_standalone, '_cache'):
            _standalone._cache[Path(phase1_output_excel).name] = str(Path(phase1_output_excel).absolute())
            if phase2_report_path.exists():
                _standalone._cache[phase2_report_name] = str(phase2_report_path.absolute())
            if validated_report_path.exists():
                _standalone._cache[validated_report_name] = str(validated_report_path.absolute())
            if final_report_path.exists() and final_report_path != validated_report_path:
                _standalone._cache[final_report_name] = str(final_report_path.absolute())

    # ── Build merged result — Phase 4/3 is the primary download ────────────────
    merged_result = result_data.copy()

    # Primary frontend output = Phase 4 (or Phase 3) Excel
    merged_result["output_file"]       = final_report_name
    merged_result["excel_file"]        = final_report_name
    merged_result["excel_url"]         = f"/api/download/{final_report_name}"
    merged_result["final_report_path"] = str(final_report_path.absolute())

    # All Excels exposed individually for the UI
    merged_result["phase1_invoice_excel"]         = Path(phase1_output_excel).name
    merged_result["phase1_invoice_excel_url"]     = f"/api/download/{Path(phase1_output_excel).name}"
    merged_result["phase2_filled_census"]         = phase2_report_name
    merged_result["phase2_filled_census_url"]     = f"/api/download/{phase2_report_name}"
    merged_result["phase3_validated_census"]      = validated_report_name
    merged_result["phase3_validated_census_url"]  = f"/api/download/{validated_report_name}"
    
    if final_report_path != validated_report_path:
        merged_result["phase4_llm_census"]      = final_report_name
        merged_result["phase4_llm_census_url"]  = f"/api/download/{final_report_name}"

    print("\n--- Flow Completed Successfully! ---")
    print(f"  Excel 1 — Invoice Extraction : {Path(phase1_output_excel).name}")
    print(f"  Excel 2 — Filled Census      : {phase2_report_name}")
    print(f"  Excel 3 — Validated Census   : {validated_report_name}")
    if final_report_path != validated_report_path:
        print(f"  Excel 4 — LLM Resolved       : {final_report_name}")
    return merged_result


if __name__ == "__main__":
    import asyncio
    if len(sys.argv) < 3:
        print("Usage: python flow_orchestrator.py <pdf_path> <template_excel> [ref_census_excel]")
    else:
        pdf = sys.argv[1]
        template = sys.argv[2]
        ref = sys.argv[3] if len(sys.argv) > 3 else None
        asyncio.run(run_flow(pdf, template, ref))