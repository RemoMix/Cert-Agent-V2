
#!/usr/bin/env python3
"""
اختبار سريع لاستخراج النص من PDF
"""

import sys
import os
import json

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from Agents.PDFtoImageAgent import PDFtoImageAgent
from Agents.JSONExtractLotAgent import JSONExtractLotAgent


def test_single_pdf(pdf_path):
    """اختبار ملف PDF واحد"""
    print(f"\\n{'='*60}")
    print(f"Testing: {os.path.basename(pdf_path)}")
    print(f"{'='*60}")
    
    json_path = f"test_output/json/{os.path.splitext(os.path.basename(pdf_path))[0]}_ocr.json"
    
    # لو الـ JSON مش موجود، اعمل Stage 1
    if not os.path.exists(json_path):
        print("\\nStage 1: PDF → Image → OCR → JSON")
        agent = PDFtoImageAgent()
        result = agent.process_pdf(pdf_path, output_dir="test_output")
        
        if not result:
            print("❌ Stage 1 failed")
            return
        
        json_path = result['json_file']
        print(f"✓ Stage 1 complete: {json_path}")
    else:
        print(f"\\nUsing existing JSON: {json_path}")
        print("(امسح الملف لو عايز تعمل OCR من جديد)")
    
    # Stage 2: JSON → Extract Lot
    print("\\nStage 2: JSON → Extract Lot")
    extract_agent = JSONExtractLotAgent()
    extraction = extract_agent.process_json(json_path)
    
    if not extraction:
        print("❌ Stage 2 failed - No lot extracted")
        
        # DEBUG: اعرض النص من JSON
        print("\\nDEBUG: النص اللي في JSON:")
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
            text = data[0]['ocr_text']
            print(text[:1000])
        return
    
    print(f"\\n✓✓✓ SUCCESS ✓✓✓")
    print(f"  Lot Numbers: {extraction['lot_numbers']}")
    print(f"  Cert Number: {extraction['certification_number']}")
    print(f"  Product: {extraction['product_name']}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        # اختبار كل الـ PDFs في Cert_Inbox
        inbox = "GetCertAgent/Cert_Inbox"
        if os.path.exists(inbox):
            pdfs = [f for f in os.listdir(inbox) if f.endswith('.pdf')]
            for pdf in pdfs:
                test_single_pdf(os.path.join(inbox, pdf))
        else:
            print(f"Inbox not found: {inbox}")
            print("Usage: python test_extraction.py <pdf_path>")
    else:
        test_single_pdf(sys.argv[1])
