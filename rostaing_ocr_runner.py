#!/usr/bin/env python3
"""
Rostaing OCR Runner

This script verifies whether rostaing-ocr and Tesseract are available, then
attempts to extract structured text from a PDF using rostaing-ocr.
If rostaing-ocr is not available, it falls back to RPVE's standard PDF text
extraction path for comparison.
"""

from __future__ import annotations

import shutil
import sys
from pathlib import Path


def check_tesseract() -> bool:
    """Return True if Tesseract is available on PATH."""
    path = shutil.which("tesseract")
    if path:
        print(f"[CHECK] Tesseract found: {path}")
        return True
    print("[CHECK] Tesseract not found in PATH. OCR fallback may fail.")
    return False


def check_rostaing() -> bool:
    """Return True if rostaing-ocr is installed."""
    try:
        import rostaing_ocr  # noqa: F401
        print("[CHECK] rostaing-ocr package is installed.")
        return True
    except ImportError:
        print("[CHECK] rostaing-ocr package is not installed.")
        return False


def save_text(text: str, output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(text, encoding="utf-8")
    print(f"[OUTPUT] Saved text to: {output_path}")


def run_rostaing_extraction(pdf_path: Path) -> Path | None:
    """Run rostaing-ocr extraction and save the structured text."""
    try:
        from schema_ocr import SchemaOCRExtractor
    except Exception as exc:
        print(f"[ERROR] Failed to import schema_ocr: {exc}")
        return None

    try:
        extractor = SchemaOCRExtractor(pdf_path)
        text = extractor.extract_layout_text(save_debug_output=True)
        if not text or not text.strip():
            print("[ERROR] rostaing-ocr returned empty text.")
            return None

        output_path = pdf_path.with_suffix(".rostaing_layout.txt")
        save_text(text, output_path)
        return output_path
    except Exception as exc:
        print(f"[ERROR] rostaing-ocr extraction failed: {exc}")
        return None


def run_standard_extraction(pdf_path: Path) -> Path | None:
    """Run the standard RPVE extraction path as a fallback."""
    try:
        from RPVE_standalone import extract_text
    except Exception as exc:
        print(f"[ERROR] Failed to import RPVE_standalone.extract_text: {exc}")
        return None

    try:
        text = extract_text(pdf_path)
        if not text or not text.strip():
            print("[ERROR] Standard RPVE extraction returned empty text.")
            return None

        output_path = pdf_path.with_suffix(".rpve_extracted.txt")
        save_text(text, output_path)
        return output_path
    except Exception as exc:
        print(f"[ERROR] Standard extraction failed: {exc}")
        return None


def main() -> int:
    if len(sys.argv) != 2:
        print("Usage: python rostaing_ocr_runner.py <path_to_pdf>")
        return 1

    pdf_path = Path(sys.argv[1])
    if not pdf_path.exists():
        print(f"[ERROR] PDF file not found: {pdf_path}")
        return 1

    print(f"[INFO] Running OCR readiness check for: {pdf_path.name}")
    tesseract_ok = check_tesseract()
    rostaing_ok = check_rostaing()

    result_path = None
    if rostaing_ok:
        print("[INFO] Attempting structured extraction with rostaing-ocr...")
        result_path = run_rostaing_extraction(pdf_path)
        if result_path:
            print("[INFO] rostaing-ocr extraction completed successfully.")
        else:
            print("[WARN] rostaing-ocr extraction failed or produced no output.")

    if not result_path:
        print("[INFO] Falling back to standard RPVE text extraction.")
        result_path = run_standard_extraction(pdf_path)
        if result_path:
            print("[INFO] Standard extraction completed successfully.")
        else:
            print("[ERROR] Both rostaing-ocr and standard extraction failed.")
            return 1

    print("[DONE] OCR extraction finished.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
