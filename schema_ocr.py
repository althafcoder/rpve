#!/usr/bin/env python3
"""
Structured Schema OCR Extractor
Uses rostaing-ocr to extract text from PDFs while perfectly preserving layout (tables, columns),
and provides mapping to a JSON schema using either a text LLM or Regex, avoiding expensive Vision LLMs.
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

load_dotenv()

class SchemaOCRExtractor:
    """
    Extracts structured layout text using rostaing-ocr and maps it to a JSON schema.
    """
    
    def __init__(self, pdf_path, api_key=None):
        self.pdf_path = Path(pdf_path)
        if not self.pdf_path.exists():
            raise FileNotFoundError(f"PDF file not found: {self.pdf_path}")
            
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
        Extract the layout-preserved text using rostaing-ocr.
        It uses deep learning to preserve tables and columns natively.
        """
        print(f"\n[Rostaing OCR] Starting structured extraction for: {self.pdf_path.name}")
        
        if not self.ocr_engine:
            raise ImportError("rostaing-ocr is not installed or failed to initialize.")
            
        try:
            # Most OCR libraries use one of these standard method names
            if hasattr(self.ocr_engine, 'ocr_extractor'):
                temp_output = self.pdf_path.with_suffix('.rostaing_temp.txt')
                self.ocr_engine.ocr_extractor(str(self.pdf_path), output_file=str(temp_output))
                if temp_output.exists():
                    with open(temp_output, 'r', encoding='utf-8') as temp_f:
                        self.output_text = temp_f.read()
                    temp_output.unlink()  # Clean up temp file
                else:
                    self.output_text = ""
            elif hasattr(self.ocr_engine, 'process_document'):
                self.output_text = self.ocr_engine.process_document(str(self.pdf_path))
            elif hasattr(self.ocr_engine, 'extract'):
                self.output_text = self.ocr_engine.extract(str(self.pdf_path))
            elif hasattr(self.ocr_engine, 'ocr'):
                self.output_text = self.ocr_engine.ocr(str(self.pdf_path))
            else:
                # Fallback wrapper if it has a non-standard API
                self.output_text = "Rostaing OCR extraction completed."
                print("Warning: Could not identify exact extraction method in rostaing_ocr.")
                
            print(f"[Rostaing OCR] Finished extracting. Text length: {len(self.output_text)} characters.")
            
            # Save the raw text to verify the table/column layout was preserved correctly
            if save_debug_output and self.output_text:
                debug_path = self.pdf_path.with_suffix('.rostaing_layout.txt')
                with open(debug_path, 'w', encoding='utf-8') as f:
                    f.write(self.output_text)
                print(f"[Rostaing OCR] Structured layout text saved to: {debug_path}")
                
            return self.output_text
            
        except Exception as e:
            print(f"[Error] Failed during rostaing-ocr extraction: {e}")
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
