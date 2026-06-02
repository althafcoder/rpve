#!/usr/bin/env python3
"""
schema_ocr_fallback.py  —  DocTR-based OCR Fallback Extractor
==============================================================
Used automatically by schema_ocr.py when rostaing-ocr fails.
Can also be run standalone:

    python schema_ocr_fallback.py <path_to_pdf>

Fallback chain (in order of preference):
  1. DocTR (doctr)      — deep-learning layout-preserving OCR
  2. pdfplumber         — reliable digital-PDF text extraction
  3. rostaing-ocr       — simple ocr_extractor() path (different from schema_ocr.py's deep-learning path)
  4. pdfminer           — older but dependency-free last resort
  5. Empty string       — graceful no-crash safety net
"""

import sys
import logging
from pathlib import Path

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Level 1: DocTR  (deep-learning, layout-preserving)
# ---------------------------------------------------------------------------
def _extract_with_doctr(pdf_path: str) -> str:
    """
    Use doctr to OCR a PDF.  Renders each page at 2x resolution to a PNG,
    runs DBNet + CRNN, then reconstructs the visual table layout using the
    same overlap-clustering algorithm as rostaing-ocr.
    """
    from doctr.io import DocumentFile
    from doctr.models import ocr_predictor
    import fitz  # PyMuPDF

    logger.info(f"[DocTR Fallback] Loading OCR predictor...")
    model = ocr_predictor(det_arch="db_resnet50", reco_arch="crnn_vgg16_bn", pretrained=True)

    try:
        import torch
        if torch.cuda.is_available():
            model.cuda()
            logger.info("[DocTR Fallback] Running on GPU (CUDA)")
        else:
            logger.info("[DocTR Fallback] Running on CPU")
    except ImportError:
        logger.info("[DocTR Fallback] torch not found — running on CPU")

    doc_fitz = fitz.open(pdf_path)
    mat = fitz.Matrix(2, 2)  # 2× zoom for higher-res images → better OCR
    full_text_pages = []

    for page_idx, page in enumerate(doc_fitz):
        logger.info(f"[DocTR Fallback] Processing page {page_idx + 1}/{len(doc_fitz)}")
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img_bytes = pix.tobytes("png")

        doc_img = DocumentFile.from_images(img_bytes)
        result  = model(doc_img)
        doctr_page = result.pages[0]

        # Flatten every word with its bounding-box geometry
        all_words = []
        for block in doctr_page.blocks:
            for line in block.lines:
                for word in line.words:
                    xmin, ymin = word.geometry[0]
                    xmax, ymax = word.geometry[1]
                    all_words.append({
                        "text":  word.value,
                        "xmin": float(xmin), "xmax": float(xmax),
                        "ymin": float(ymin), "ymax": float(ymax),
                    })

        if not all_words:
            full_text_pages.append(f"--- Page {page_idx + 1} ---\n")
            continue

        # Sort top-to-bottom, then cluster into visual lines using 40% overlap
        all_words.sort(key=lambda w: w["ymin"])
        lines = []
        for word in all_words:
            placed = False
            for line in lines:
                ly_min = sum(w["ymin"] for w in line) / len(line)
                ly_max = sum(w["ymax"] for w in line) / len(line)
                overlap = max(0, min(word["ymax"], ly_max) - max(word["ymin"], ly_min))
                wh = word["ymax"] - word["ymin"]
                lh = ly_max - ly_min
                if overlap > 0.4 * min(wh, lh):
                    line.append(word)
                    placed = True
                    break
            if not placed:
                lines.append([word])

        lines.sort(key=lambda ln: sum(w["ymin"] for w in ln) / len(ln))

        # Reconstruct visual table layout with tab-stop gaps
        output_lines = []
        for line in lines:
            line.sort(key=lambda w: w["xmin"])
            text = ""
            last_x = 0.0
            for word in line:
                if last_x == 0.0:
                    text += word["text"]
                else:
                    gap = word["xmin"] - last_x
                    text += ("   \t   " if gap > 0.1 else " ") + word["text"]
                last_x = word["xmax"]
            output_lines.append(text)

        full_text_pages.append(f"--- Page {page_idx + 1} ---\n" + "\n".join(output_lines))

    doc_fitz.close()
    return "\n\n".join(full_text_pages)


# ---------------------------------------------------------------------------
# Level 2: pdfplumber  (digital PDFs only — fast, no ML needed)
# ---------------------------------------------------------------------------
def _extract_with_pdfplumber(pdf_path: str) -> str:
    import pdfplumber
    logger.info("[pdfplumber Fallback] Extracting text from digital PDF...")
    pages = []
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            pages.append(f"--- Page {i} ---\n{text}")
    return "\n\n".join(pages)


# ---------------------------------------------------------------------------
# Level 3: rostaing-ocr  (simple ocr_extractor path — no deep-learning overhead)
# ---------------------------------------------------------------------------
def _extract_with_rostaing(pdf_path: str) -> str:
    """
    Call rostaing-ocr's built-in ocr_extractor() helper.
    This is the lightweight path (NOT the DocTR deep-learning path used by
    schema_ocr.py), so it provides genuinely distinct coverage as a fallback.
    """
    import rostaing_ocr  # ImportError bubbles up → next level is tried
    logger.info("[rostaing-ocr Fallback] Attempting simple ocr_extractor() extraction...")

    import tempfile, os
    with tempfile.NamedTemporaryFile(suffix=".txt", delete=False) as tmp:
        tmp_path = tmp.name

    try:
        engine = rostaing_ocr.RostaingOCR() if hasattr(rostaing_ocr, "RostaingOCR") else rostaing_ocr
        if hasattr(engine, "ocr_extractor"):
            engine.ocr_extractor(pdf_path, output_file=tmp_path)
            with open(tmp_path, "r", encoding="utf-8") as f:
                text = f.read()
            logger.info(f"[rostaing-ocr Fallback] Extracted {len(text)} characters.")
            return text
        else:
            raise AttributeError("rostaing_ocr has no ocr_extractor() method")
    finally:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)


# ---------------------------------------------------------------------------
# Level 4: pdfminer  (older but zero-dependency last resort)
# ---------------------------------------------------------------------------
def _extract_with_pdfminer(pdf_path: str) -> str:
    from pdfminer.high_level import extract_text
    logger.info("[pdfminer Fallback] Extracting text via pdfminer...")
    return extract_text(pdf_path)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------
class DoctrOCRFallback:
    """
    Drop-in fallback for schema_ocr.py's SchemaOCRExtractor.
    Tries four escalating extraction strategies so something always returns.

    Fallback chain:
      Level 1 — DocTR         (deep-learning, layout-preserving)
      Level 2 — pdfplumber    (fast digital-PDF extraction)
      Level 3 — rostaing-ocr  (simple ocr_extractor() path)
      Level 4 — pdfminer      (dependency-free last resort)
      Level 5 — ""            (never crashes)

    Usage (from schema_ocr.py):
        from schema_ocr_fallback import DoctrOCRFallback
        text = DoctrOCRFallback(pdf_path).extract()
    """

    def __init__(self, pdf_path: str):
        self.pdf_path = str(pdf_path)

    def extract(self) -> str:
        """
        Run the fallback chain and return extracted text.
        Never raises — returns an empty string on total failure.
        """
        # --- Level 1: DocTR ---
        try:
            text = _extract_with_doctr(self.pdf_path)
            if text.strip():
                logger.info("[DocTR Fallback] Extraction succeeded.")
                return text
        except Exception as e:
            logger.warning(f"[DocTR Fallback] Failed: {e}")

        # --- Level 2: pdfplumber ---
        try:
            text = _extract_with_pdfplumber(self.pdf_path)
            if text.strip():
                logger.info("[pdfplumber Fallback] Extraction succeeded.")
                return text
        except Exception as e:
            logger.warning(f"[pdfplumber Fallback] Failed: {e}")

        # --- Level 3: rostaing-ocr (simple path) ---
        try:
            text = _extract_with_rostaing(self.pdf_path)
            if text.strip():
                logger.info("[rostaing-ocr Fallback] Extraction succeeded.")
                return text
        except Exception as e:
            logger.warning(f"[rostaing-ocr Fallback] Failed: {e}")

        # --- Level 4: pdfminer ---
        try:
            text = _extract_with_pdfminer(self.pdf_path)
            if text.strip():
                logger.info("[pdfminer Fallback] Extraction succeeded.")
                return text
        except Exception as e:
            logger.warning(f"[pdfminer Fallback] Failed: {e}")

        # Total failure — return empty so callers don't crash
        logger.error("[DoctrOCRFallback] All extraction methods failed. Returning empty string.")
        return ""


# ---------------------------------------------------------------------------
# Standalone CLI (mirrors the original 10-line script behaviour)
# ---------------------------------------------------------------------------
def main():
    logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

    if len(sys.argv) < 2:
        print(f"Usage: python {Path(__file__).name} <path_to_pdf>  [output.txt]")
        sys.exit(1)

    pdf_path   = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else "output.txt"

    fallback = DoctrOCRFallback(pdf_path)
    text = fallback.extract()

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(text)

    print(f"Extracted text saved to: {output_path}  ({len(text)} characters)")


if __name__ == "__main__":
    main()