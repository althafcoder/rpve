"""
flow_orchestrator.py
====================
Synchronous RPVE pipeline orchestrator.

run_job() is called from a worker thread (job_worker.py).
Because it runs in a thread (not on the asyncio event loop),
all subprocess.run() calls here are safe — they block only
the current thread, not the entire server.

Business logic is 100% preserved from the original run_flow():
  - .xls → .xlsx conversion
  - 3-way swap intelligence (name / volume / format)
  - Phase 2 template classification (type1 / type2 / type3)
  - Phase 3 data validation (non-fatal fallback)
  - Phase 4 LLM resolution (non-fatal fallback)
"""

import logging
import sys
import subprocess
from pathlib import Path
from typing import Callable, Optional

import pandas as pd

# Set up paths so that RPVE_standalone can be imported regardless of cwd
CURRENT_DIR = Path(__file__).parent.absolute()
sys.path.append(str(CURRENT_DIR))

try:
    import RPVE_standalone
    from RPVE_standalone import process_invoice_data_sync, OUTPUT_DIR
except ImportError as e:
    print(f"Error: Could not import RPVE_standalone.py. {e}")
    sys.exit(1)

# ─────────────────────────────────────────────────────────────────────────────
# Helpers (preserved from original)
# ─────────────────────────────────────────────────────────────────────────────

def ensure_xlsx(file_path: Optional[str]) -> Optional[str]:
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
        import warnings
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            df = pd.read_excel(excel_path, nrows=25, header=None)
        all_text = " ".join([str(val).lower() for val in df.values.flatten() if pd.notna(val)])
        
        # Type 1 Detection (Engage/Kaiser)
        if "ee row" in all_text or "relation-ship to employee" in all_text or "kaiser networks" in all_text:
            return "type1"
            
        # Type 3 Detection (RAPT Blue Headers)
        if "data row" in all_text or "cobra participant" in all_text or "discrepancies" in all_text:
            return "type3"
            
        # Type 2 (Basic Titan Intake / Generic Census)
        elif "first name" in all_text or "last name" in all_text or "name" in all_text:
            return "type2"

        # Default fallback to type2 (dynamic column header mapping)
        return "type2"
    except Exception as e:
        print(f"Error reading Excel: {e}")
        return "type2"


# ─────────────────────────────────────────────────────────────────────────────
# Main entry point — called from job_worker._execute_job()
# ─────────────────────────────────────────────────────────────────────────────

def run_job(
    job_id: str,
    pdf_path: Path,
    template_path: str,
    ref_census_path: Optional[str],
    job_dir: Path,
    status_callback: Callable[[str, Optional[str]], None],
    logger: logging.Logger,
) -> dict:
    """
    Run the full RPVE pipeline for a single job.

    Parameters
    ----------
    job_id          : Unique job identifier (backend-internal only)
    pdf_path        : Absolute Path to the uploaded invoice PDF/XLSX/CSV
    template_path   : Absolute path string to the census template Excel
    ref_census_path : Optional absolute path string to a reference census Excel
    job_dir         : Per-job workspace root (jobs/{job_id}/)
    status_callback : Called at each phase transition → (phase_name, template_type?)
    logger          : Per-job file logger

    Returns
    -------
    dict — the merged result dict (same shape as the old run_flow() return value)
    """

    work_dir   = job_dir / "work"
    output_dir = job_dir / "output"
    work_dir.mkdir(exist_ok=True)
    output_dir.mkdir(exist_ok=True)

    logger.info("--- Starting End-to-End RPVE Flow (job_id=%s) ---", job_id)

    # ── PHASE 1: PDF Extraction ──────────────────────────────────────────────
    status_callback("phase1")
    logger.info("[1] Running Phase 1 (Extraction) on: %s", pdf_path.name)
    try:
        result_data = process_invoice_data_sync(pdf_path, pdf_path.name, out_dir=work_dir)
        phase1_output_excel = result_data["excel_path"]
        logger.info("    -> Phase 1 Success! Extracted: %s", Path(phase1_output_excel).name)

        phase2_report_path    = work_dir / f"FINAL_AUDIT_REPORT_{job_id}.xlsx"
        validated_report_path = work_dir / f"VALIDATED_AUDIT_REPORT_{job_id}.xlsx"

    except Exception as e:
        import traceback
        logger.error("    -> Phase 1 Failed: %s\n%s", e, traceback.format_exc())
        raise

    # ── PHASE 1.5: Pre-process Templates (.xls -> .xlsx) ────────────────────
    template_path   = ensure_xlsx(template_path)
    ref_census_path = ensure_xlsx(ref_census_path)

    # ── PHASE 2: Template Classification + Fill ──────────────────────────────
    status_callback("classifying")
    logger.info("[2] Analyzing Template: %s", template_path)

    template_type = classify_excel_template(Path(template_path))

    if ref_census_path:
        ref_type = classify_excel_template(Path(ref_census_path))

        # ── ROBUST SWAP INTELLIGENCE (preserved exactly from original) ───────
        t_name = Path(template_path).name.lower()
        r_name = Path(ref_census_path).name.lower()

        logger.info("    -> Swap Check: Template=%s, Ref=%s", t_name, r_name)
        logger.info("    -> Type Check: Template=%s, Ref=%s", template_type, ref_type)

        swapped = False

        # 1. Name-based hint
        if template_type not in ["type1", "type3"] and (
            ("rapt" in r_name and "rapt" not in t_name)
            or ("template" in r_name and "template" not in t_name)
        ):
            logger.info("    -> Name-based Swap: User put RAPT/Template file in Reference slot.")
            template_path, ref_census_path = ref_census_path, template_path
            template_type = classify_excel_template(Path(template_path))
            ref_type      = classify_excel_template(Path(ref_census_path))
            swapped = True
            logger.info("    -> Post-Swap Types: Template=%s, Ref=%s", template_type, ref_type)

        # 2. Volume-based check
        if not swapped:
            try:
                import warnings
                with warnings.catch_warnings():
                    warnings.simplefilter("ignore")
                    t_df = pd.read_excel(template_path, nrows=100)
                    r_df = pd.read_excel(ref_census_path, nrows=100)
                    t_val = t_df.iloc[:, :3].count().sum()
                    r_val = r_df.iloc[:, :3].count().sum()
                    logger.info("    -> Content Check: Template=%s values, Ref=%s values", t_val, r_val)
                    if t_val > (r_val + 10):
                        logger.info("    -> Volume-based Swap Triggered!")
                        template_path, ref_census_path = ref_census_path, template_path
                        template_type = classify_excel_template(Path(template_path))
                        ref_type      = classify_excel_template(Path(ref_census_path))
                        swapped = True
            except Exception as e:
                logger.warning("    -> Volume check failed: %s", e)

        # 3. Format-based fallback
        if not swapped and ref_type == "type3" and template_type != "type3":
            logger.info("    -> Format-based Swap: User put RAPT file in Reference slot.")
            template_path, ref_census_path = ref_census_path, template_path
            template_type = "type3"
        elif ref_census_path:
            template_type = "type3"

        logger.info("    -> Final Decision: Template Type=%s", template_type.upper())
    else:
        logger.info("    -> Detected Template Type: %s", template_type.upper())

    # Record template_type in DB
    status_callback("phase2", template_type)

    phase2_base = CURRENT_DIR / "phase_2" / "template_fill"

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
                raise ValueError("Type 3 requires a Reference Census file!")
        script = phase2_base / "type3" / "fill_template.py"
        cmd = [sys.executable, str(script), phase1_output_excel, ref_census_path, template_path, str(phase2_report_path)]

    else:
        raise ValueError(f"Could not identify Excel template type: {template_type}")

    logger.info("[3] Executing Phase 2 (Template Fill)...")
    logger.info("    Running command: %s", " ".join(cmd))
    result = subprocess.run(cmd, capture_output=True, text=True, cwd=str(script.parent))

    if result.returncode != 0:
        logger.error("--- Flow Failed in Phase 2 ---\n%s", result.stderr)
        raise RuntimeError(f"Phase 2 failed: {result.stderr}")

    logger.info("    -> Phase 2 (Census Fill) Success!")
    if result.stdout:
        logger.debug(result.stdout)

    # ── PHASE 3: Data Validation ─────────────────────────────────────────────
    status_callback("phase3")
    logger.info("[4] Executing Phase 3 (Data Validation)...")
    validation_script = phase2_base / "data_validation.py"

    if not validation_script.exists():
        logger.warning("    [WARN] data_validation.py not found. Skipping Phase 3.")
        validated_report_path = phase2_report_path
    else:
        val_cmd = [
            sys.executable, str(validation_script),
            str(phase2_report_path),
            phase1_output_excel,
            str(validated_report_path),
            "--threshold", "85",
            "--template-type", template_type,
        ]
        logger.info("    Running command: %s", " ".join(val_cmd))
        val_result = subprocess.run(val_cmd, capture_output=True, text=True)

        if val_result.returncode == 0:
            logger.info("    -> Phase 3 (Data Validation) Success!")
            if val_result.stdout:
                logger.debug(val_result.stdout)
        else:
            logger.warning("    [WARN] Phase 3 failed (non-fatal). Using Phase 2 output.\n%s", val_result.stderr)
            validated_report_path = phase2_report_path

    # ── PHASE 4: LLM Resolution ───────────────────────────────────────────────
    status_callback("phase4")
    logger.info("[5] Executing Phase 4 (LLM Resolution)...")
    llm_script     = phase2_base / "llm_resolution.py"
    audit_json     = validated_report_path.with_suffix(".audit.json")
    llm_report_path = validated_report_path.with_name(
        validated_report_path.name.replace("VALIDATED_", "LLM_RESOLVED_")
    )

    if not llm_script.exists() or not audit_json.exists():
        logger.warning("    [WARN] LLM resolution requirements not met. Skipping Phase 4.")
        final_report_path = validated_report_path
    else:
        llm_cmd = [
            sys.executable, str(llm_script),
            str(validated_report_path),
            str(audit_json),
            "--output", str(llm_report_path),
            "--template-type", template_type,
        ]
        logger.info("    Running command: %s", " ".join(llm_cmd))
        llm_result = subprocess.run(llm_cmd, capture_output=True, text=True)

        if llm_result.returncode == 0 and llm_report_path.exists():
            logger.info("    -> Phase 4 (LLM Resolution) Success!")
            if llm_result.stdout:
                logger.debug(llm_result.stdout)
            if llm_result.stderr:
                logger.debug(llm_result.stderr)
            final_report_path = llm_report_path
        else:
            logger.warning(
                "    [WARN] Phase 4 failed or skipped. Using Phase 3 output.\n%s",
                llm_result.stderr if llm_result else "",
            )
            final_report_path = validated_report_path

    # ── Build merged result (same shape as original run_flow return value) ───
    merged_result = result_data.copy()

    final_report_name     = final_report_path.name
    phase2_report_name    = phase2_report_path.name
    validated_report_name = validated_report_path.name

    merged_result["output_file"]       = final_report_name
    merged_result["excel_file"]        = final_report_name
    merged_result["excel_url"]         = f"/api/download/{final_report_name}"
    merged_result["final_report_path"] = str(final_report_path.absolute())

    merged_result["phase1_invoice_excel"]         = Path(phase1_output_excel).name
    merged_result["phase1_invoice_excel_url"]     = f"/api/download/{Path(phase1_output_excel).name}"
    merged_result["phase2_filled_census"]         = phase2_report_name
    merged_result["phase2_filled_census_url"]     = f"/api/download/{phase2_report_name}"
    merged_result["phase3_validated_census"]      = validated_report_name
    merged_result["phase3_validated_census_url"]  = f"/api/download/{validated_report_name}"

    if final_report_path != validated_report_path:
        merged_result["phase4_llm_census"]      = final_report_name
        merged_result["phase4_llm_census_url"]  = f"/api/download/{final_report_name}"

    logger.info("--- Flow Completed Successfully! ---")
    logger.info("  Excel 1 — Invoice Extraction : %s", Path(phase1_output_excel).name)
    logger.info("  Excel 2 — Filled Census      : %s", phase2_report_name)
    logger.info("  Excel 3 — Validated Census   : %s", validated_report_name)
    if final_report_path != validated_report_path:
        logger.info("  Excel 4 — LLM Resolved       : %s", final_report_name)

    return merged_result


# ─────────────────────────────────────────────────────────────────────────────
# Legacy CLI entry point (for direct invocation testing)
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import uuid
    if len(sys.argv) < 3:
        print("Usage: python flow_orchestrator.py <pdf_path> <template_excel> [ref_census_excel]")
        sys.exit(1)

    _pdf      = Path(sys.argv[1])
    _template = sys.argv[2]
    _ref      = sys.argv[3] if len(sys.argv) > 3 else None
    _job_id   = f"cli_{uuid.uuid4().hex[:8]}"
    _job_dir  = Path("jobs") / _job_id
    _job_dir.mkdir(parents=True, exist_ok=True)

    _logger = logging.getLogger("rpve.cli")
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(message)s")

    result = run_job(
        job_id=_job_id,
        pdf_path=_pdf,
        template_path=_template,
        ref_census_path=_ref,
        job_dir=_job_dir,
        status_callback=lambda phase, ttype=None: print(f"[STATUS] {phase}"),
        logger=_logger,
    )
    import json
    print(json.dumps(result, indent=2, default=str))