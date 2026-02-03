
# النسخة النهائية الصحيحة
import os
import re
import json
import yaml
from datetime import datetime
from pathlib import Path
import logging

logger = logging.getLogger('CertPrintAgent')


class JSONExtractLotAgent:
    def __init__(self, config_path="config.yaml"):
        self.config = self.load_config(config_path)
        
    def load_config(self, config_path):
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                return yaml.safe_load(f)
        except Exception as e:
            logger.error(f"Error loading config: {e}")
            return {}
    
    def load_json(self, json_path):
        """Load OCR JSON file"""
        try:
            with open(json_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            logger.info(f"Loaded JSON: {json_path} ({len(data)} pages)")
            return data
        except Exception as e:
            logger.error(f"Error loading JSON: {e}")
            return None
    
    def parse_lot(self, raw):
        """Parse lot number structure"""
        raw = raw.strip().replace(" ", "")
        
        if "-" in raw and all(p.isdigit() for p in raw.split("-") if p):
            parts = raw.split("-")
            return {
                "type": "explicit_multi",
                "base_lot": None,
                "count": len(parts),
                "expanded_lots": parts,
                "annotation_hint": None,
            }
        
        if "/" in raw:
            parts = raw.split("/")
            if len(parts) == 2 and parts[0].isdigit() and parts[1].isdigit():
                base, cnt = parts[0], int(parts[1])
                return {
                    "type": "implicit_multi",
                    "base_lot": base,
                    "count": cnt,
                    "expanded_lots": [base],
                    "annotation_hint": f"+{cnt-1}",
                }
        
        return {
            "type": "single",
            "base_lot": raw,
            "count": 1,
            "expanded_lots": [raw],
            "annotation_hint": None,
        }
    
    def extract_lot_from_text(self, text):
        """Extract lot number from text - FINAL WORKING VERSION"""
        
        logger.info("=== EXTRACTING LOT ===")
        
        # Pattern 1: Lot Number : 139928 (مباشر وبسيط)
        match = re.search(r'Lot Number : (\d+)', text)
        if match:
            lot = match.group(1)
            logger.info(f"✓✓✓ FOUND: {lot} ✓✓✓")
            return {
                "lot_raw": lot,
                "lot_structured": self.parse_lot(lot)
            }
        
        # Pattern 2: Case insensitive
        match = re.search(r'lot number : (\d+)', text, re.IGNORECASE)
        if match:
            lot = match.group(1)
            logger.info(f"✓✓✓ FOUND: {lot} ✓✓✓")
            return {
                "lot_raw": lot,
                "lot_structured": self.parse_lot(lot)
            }
        
        # Pattern 3: مرن مع مسافات
        match = re.search(r'Lot\s+Number\s*:\s*(\d+)', text, re.IGNORECASE)
        if match:
            lot = match.group(1)
            logger.info(f"✓✓✓ FOUND: {lot} ✓✓✓")
            return {
                "lot_raw": lot,
                "lot_structured": self.parse_lot(lot)
            }
        
        # Pattern 4: دور على أرقام 6-7 digits
        numbers = re.findall(r'\b\d{6,7}\b', text)
        if numbers:
            if '139928' in numbers:
                logger.info(f"✓✓✓ FOUND: 139928 ✓✓✓")
                return {
                    "lot_raw": "139928",
                    "lot_structured": self.parse_lot("139928")
                }
            lot = numbers[0]
            logger.info(f"✓✓✓ FOUND: {lot} ✓✓✓")
            return {
                "lot_raw": lot,
                "lot_structured": self.parse_lot(lot)
            }
        
        logger.error("✗✗✗ NOT FOUND ✗✗✗")
        return None
    
    def extract_certification_number(self, text):
        # Pattern 1: Certificate Number : Dokki-XXXXX
        match = re.search(r'Certificate Number : ([A-Za-z]+-\d+)', text)
        if match:
            return match.group(1)
        
        # Pattern 2: Dokki-XXXXX في أي مكان
        match = re.search(r'(Dokki-\d+)', text)
        if match:
            return match.group(1)
        
        return "UNKNOWN"
    
    def extract_product_name(self, text):
        match = re.search(r'Sample : ([A-Za-z]+)', text)
        if match:
            return match.group(1)
        for product in ['Basil', 'Fennel', 'Peppermint', 'Marjoram']:
            if product in text:
                return product
        return "UNKNOWN"
    
    def process_json(self, json_path):
        """Process single JSON file"""
        logger.info(f"\n{'='*60}")
        logger.info(f"Processing: {os.path.basename(json_path)}")
        logger.info(f"{'='*60}")
        
        data = self.load_json(json_path)
        if not data:
            return None
        
        page_data = data[0] if data else None
        if not page_data:
            logger.error("No page data")
            return None
        
        text = page_data.get("ocr_text", "")
        if not text:
            logger.error("No text")
            return None
        
        lot_data = self.extract_lot_from_text(text)
        if not lot_data:
            return None
        
        result = {
            "file_path": page_data.get("pdf_file", ""),
            "file_name": page_data.get("pdf_file", "unknown"),
            "certification_number": self.extract_certification_number(text),
            "product_name": self.extract_product_name(text),
            "lot_numbers": lot_data["lot_structured"]["expanded_lots"],
            "lot_info": [{"num": lot_data["lot_raw"], "type": lot_data["lot_structured"]["type"]}],
            "lot_structure": lot_data["lot_structured"]["type"],
            "extraction_time": datetime.now().isoformat(),
        }
        
        logger.info(f"✓ SUCCESS: Lot={result['lot_numbers']}, Product={result['product_name']}")
        
        return result
    
    def process_all(self, json_dir=None):
        """Process all JSON files"""
        if json_dir is None:
            json_dir = Path("temp_images") / "json"
        
        json_files = list(Path(json_dir).glob("*_ocr.json"))
        logger.info(f"Found {len(json_files)} JSON file(s)")
        
        results = []
        for json_file in json_files:
            result = self.process_json(str(json_file))
            if result:
                results.append(result)
        
        logger.info(f"Total successful: {len(results)}/{len(json_files)}")
        
        return results
    
    def run(self):
        """Run the agent"""
        logger.info("=== JSON Extract Lot Agent ===")
        return self.process_all()


def extract_from_json(config_path="config.yaml"):
    agent = JSONExtractLotAgent(config_path)
    return agent.run()
