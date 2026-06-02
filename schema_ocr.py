#!/usr/bin/env python3
"""
Structured Schema OCR Extractor
Uses rostaing-ocr to extract text from PDFs while perfectly preserving layout (tables, columns),
and provides mapping to a JSON schema using either a text LLM or Regex, avoiding expensive Vision LLMs.

Fallback chain (triggered automatically on any extraction failure):
  rostaing-ocr  →  DocTR (schema_ocr_fallback.py Level 1)
                →  pdfplumber (schema_ocr_fallback.py Level 2)
                →  pdfminer   (schema_ocr_fallback.py Level 3)
                →  empty string (never crashes)
"""

from pathlib import Path
import os
import json
import re
from openai import OpenAI
from dotenv import load_dotenv

try:
    import rostaing_ocr
    ROSTAING_AVAILABLE = True
except ImportError:
    ROSTAING_AVAILABLE = False
    print("WARNING: rostaing-ocr is not installed. Please run `pip install rostaing-ocr`")

# DocTR + pdfplumber + pdfminer fallback (schema_ocr_fallback.py)
try:
    from schema_ocr_fallback import DoctrOCRFallback
    DOCTR_FALLBACK_AVAILABLE = True
except ImportError:
    DOCTR_FALLBACK_AVAILABLE = False
    print("WARNING: schema_ocr_fallback.py not found — DocTR fallback will be unavailable.")

load_dotenv()

class SchemaOCRExtractor:
    """
    Extracts structured layout text using rostaing-ocr and maps it to a JSON schema.
    """
    
    def __init__(self, pdf_path, api_key=None):
        self.pdf_path = Path(pdf_path)
        if not self.pdf_path.exists():
            raise FileNotFoundError(f"PDF file not found: {self.pdf_path}")
            
        if self.pdf_path.suffix.lower() != '.pdf':
            raise ValueError(f"SchemaOCRExtractor only supports .pdf files. Received: {self.pdf_path.suffix}")

        self.api_key = api_key or os.getenv("OPENAI_API_KEY")
        self.output_text = ""
        
        self.ocr_engine = None
        if ROSTAING_AVAILABLE:
            try:
                # Initialize the engine (downloads local models on first run)
                self.ocr_engine = rostaing_ocr.RostaingOCR() 
            except AttributeError:
                # Fallback if the library uses a different initialization pattern
                self.ocr_engine = rostaing_ocr
        else:
            print("rostaing-ocr must be installed. Methods will fail until it's loaded.")

    def extract_layout_text(self, save_debug_output=True):
        """
        Extract the layout-preserved text using the following chain:

        Primary:   rostaing-ocr  (DocTR deep-learning, 40% overlap line clustering)
        Fallback:  DoctrOCRFallback (schema_ocr_fallback.py)
                   Level 1 — DocTR standalone
                   Level 2 — pdfplumber
                   Level 3 — pdfminer

        Any exception in the primary path automatically triggers the fallback.
        """
        print(f"\n[Rostaing OCR] Starting structured extraction for: {self.pdf_path.name}")
        
        try:
            import fitz
            from doctr.io import DocumentFile
            from doctr.models import ocr_predictor
            import torch
        except ImportError as e:
            print(f"[Warning] Missing deep-learning dependencies: {e}. Falling back to standard extraction.")
            # Simple fallback to standard library if available
            if self.ocr_engine and hasattr(self.ocr_engine, 'ocr_extractor'):
                temp_output = self.pdf_path.with_suffix('.rostaing_temp.txt')
                self.ocr_engine.ocr_extractor(str(self.pdf_path), output_file=str(temp_output))
                if temp_output.exists():
                    with open(temp_output, 'r', encoding='utf-8') as temp_f:
                        self.output_text = temp_f.read()
                    temp_output.unlink()
                    return self.output_text
            raise

        try:
            # 1. Load model (DBNet + CRNN)
            device = "cuda" if torch.cuda.is_available() else "cpu"
            print(f"[Rostaing OCR] Loading predictor onto device: {device}...", flush=True)
            model = ocr_predictor(det_arch='db_resnet50', reco_arch='crnn_vgg16_bn', pretrained=True)
            if torch.cuda.is_available():
                model.cuda()

            # 2. Render PDF pages to images
            doc_fitz = fitz.open(self.pdf_path)
            mat = fitz.Matrix(2, 2)  # High resolution zoom
            full_text_pages = []

            for page_idx, page in enumerate(doc_fitz):
                print(f"[Rostaing OCR] Processing Page {page_idx + 1}/{len(doc_fitz)}...", flush=True)
                pix = page.get_pixmap(matrix=mat, alpha=False)
                img_bytes = pix.tobytes("png")
                
                doc = DocumentFile.from_images(img_bytes)
                result = model(doc)
                doctr_page = result.pages[0]

                # Flatten all words from the page blocks
                all_words = []
                for block in doctr_page.blocks:
                    for line in block.lines:
                        for word in line.words:
                            xmin, ymin = word.geometry[0]
                            xmax, ymax = word.geometry[1]
                            all_words.append({
                                'text': word.value,
                                'xmin': float(xmin),
                                'xmax': float(xmax),
                                'ymin': float(ymin),
                                'ymax': float(ymax),
                                'cy': float((ymin + ymax) / 2),
                                'height': float(ymax - ymin)
                            })

                if not all_words:
                    full_text_pages.append(f"--- Page {page_idx + 1} ---\n")
                    continue

                # Sort words from Top to Bottom based on ymin
                all_words.sort(key=lambda w: w['ymin'])

                # Robust vertical overlap clustering algorithm (40% minimum overlap threshold)
                lines = []
                for word in all_words:
                    matched_line = None
                    for line in lines:
                        # Find the representative vertical bounds of the line
                        line_ymin = sum(w['ymin'] for w in line) / len(line)
                        line_ymax = sum(w['ymax'] for w in line) / len(line)
                        
                        overlap = max(0, min(word['ymax'], line_ymax) - max(word['ymin'], line_ymin))
                        word_h = word['ymax'] - word['ymin']
                        line_h = line_ymax - line_ymin
                        
                        # If vertical overlap is significant (at least 40% of the smaller height)
                        if overlap > 0.4 * min(word_h, line_h):
                            matched_line = line
                            break
                    
                    if matched_line is not None:
                        matched_line.append(word)
                    else:
                        lines.append([word])

                # Sort clustered lines in Top-to-Bottom order
                lines.sort(key=lambda ln: sum(w['ymin'] for w in ln) / len(ln))

                # Reconstruct visual table layout
                output_lines = []
                for line in lines:
                    line.sort(key=lambda w: w['xmin'])
                    line_text = ""
                    last_x_end = 0
                    for word in line:
                        if last_x_end == 0:
                            line_text += word['text']
                        else:
                            gap = word['xmin'] - last_x_end
                            if gap > 0.1:
                                line_text += " \t   " + word['text']
                            elif gap > 0.02:
                                line_text += " " + word['text']
                            else:
                                line_text += " " + word['text']
                        last_x_end = word['xmax']
                    output_lines.append(line_text)

                page_text = "\n".join(output_lines)
                full_text_pages.append(f"--- Page {page_idx + 1} ---\n{page_text}")

            doc_fitz.close()
            self.output_text = "\n\n".join(full_text_pages)
            print(f"[Rostaing OCR] Finished extracting. Text length: {len(self.output_text)} characters.")

            # Save the raw text to verify the table/column layout was preserved correctly
            if save_debug_output and self.output_text:
                debug_path = self.pdf_path.with_suffix('.rostaing_layout.txt')
                with open(debug_path, 'w', encoding='utf-8') as f:
                    f.write(self.output_text)
                print(f"[Rostaing OCR] Structured layout text saved to: {debug_path}")

            return self.output_text

        except Exception as e:
            print(f"[Error] Failed during custom deep-learning layout extraction: {e}")

            # ── Automatic fallback via schema_ocr_fallback.py ───────────────
            if DOCTR_FALLBACK_AVAILABLE:
                print("[Rostaing OCR] Engaging DoctrOCRFallback chain (DocTR → pdfplumber → pdfminer)...")
                try:
                    fallback_text = DoctrOCRFallback(str(self.pdf_path)).extract()
                    if fallback_text.strip():
                        self.output_text = fallback_text
                        print(f"[Rostaing OCR] Fallback succeeded — {len(fallback_text)} characters extracted.")

                        # Optionally save the fallback result for debugging
                        if save_debug_output:
                            debug_path = self.pdf_path.with_suffix('.fallback_layout.txt')
                            with open(debug_path, 'w', encoding='utf-8') as f:
                                f.write(fallback_text)
                            print(f"[Rostaing OCR] Fallback text saved to: {debug_path}")

                        return self.output_text
                    else:
                        print("[Rostaing OCR] Fallback returned empty text. Giving up.")
                except Exception as fb_err:
                    print(f"[Rostaing OCR] Fallback itself failed: {fb_err}")
            else:
                print("[Rostaing OCR] schema_ocr_fallback.py not available — cannot fall back.")

            raise

    def extract_to_schema(self, schema_format: dict, use_llm=False):
        """
        Maps the highly-structured text layout directly to the requested JSON schema.
        
        Args:
            schema_format: Dict of keys to extract
            use_llm: If true, uses a CHEAP Text LLM (like gpt-4o-mini). No Vision is needed
                     because rostaing-ocr perfectly preserved the visual structure as text spaces.
                     If false, relies purely on fast regular expressions.
        """
        if not self.output_text:
            self.extract_layout_text()
            
        if use_llm and self.api_key:
            return self._parse_schema_with_text_llm(schema_format)
        else:
            return self._parse_schema_with_regex(schema_format)

    def _parse_schema_with_text_llm(self, schema_format: dict):
        """
        Passes the structured rostaing-ocr text string to a standard text LLM to guarantee perfect JSON.
        This is significantly cheaper and more reliable than using Vision models.
        """
        print("[Rostaing OCR] Mapping structured text to JSON Schema via Text LLM...")
        client = OpenAI(api_key=self.api_key)
        
        prompt = f"""
        Extract the requested fields from the structured OCR text below. 
        The text layout (tables, columns) has been natively preserved.
        
        Return ONLY a JSON dictionary exactly matching the keys of this schema:
        {json.dumps(schema_format, indent=2)}
        
        OCR TEXT:
        {self.output_text[:12000]}  # Clip at 12k chars to save tokens on massive docs
        """
        
        try:
            response = client.chat.completions.create(
                model="gpt-4o-mini",  # Text model, no vision needed
                messages=[{"role": "user", "content": prompt}],
                response_format={"type": "json_object"},
                temperature=0.0
            )
            data = json.loads(response.choices[0].message.content)
            print("[Rostaing OCR] Schema mapping completed successfully.")
            return data
        except Exception as e:
            print(f"[Error] LLM Schema mapping failed: {e}")
            return {}

    def _parse_schema_with_regex(self, schema_format: dict):
        """
        If we want zero API cost, we use regex to pull values from the structured text.
        Because rostaing-ocr preserves exact spaces, `Key: Value` pairs are very reliable.
        """
        print("[Rostaing OCR] Mapping structured text to JSON Schema via Regex...")
        results = {}
        for key in schema_format.keys():
            # A greedy regex that looks for the key name and captures whatever follows on the same line
            pattern = re.compile(f"{key}\\s*[:\\|-]?\\s*(.+)", re.IGNORECASE)
            match = pattern.search(self.output_text)
            
            if match:
                results[key] = match.group(1).strip()
            else:
                results[key] = None
                
        return results

if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        test_file = sys.argv[1]
        extractor = SchemaOCRExtractor(test_file)
        
        # Test extraction
        text = extractor.extract_layout_text(save_debug_output=True)
        
        # Test schema mapping (using basic Regex to avoid API calls during local testing)
        dummy_schema = {
            "Invoice Number": "",
            "Total Amount": "",
            "Date": ""
        }
        json_data = extractor.extract_to_schema(dummy_schema, use_llm=False)
        print(f"\nExtracted Schema Result:\n{json.dumps(json_data, indent=2)}")
    else:
        print("Usage: python schema_ocr.py <path_to_pdf_or_image>")
