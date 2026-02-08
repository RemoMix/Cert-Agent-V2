# ExtractLotAgent.py - استخراج رقم اللوت من اسم الملف مع دعم كل الأنماط
import os
import re
import yaml
from datetime import datetime
import logging

logger = logging.getLogger('CertPrintAgent')

class ExtractLotAgent:
    def __init__(self, config_path="config.yaml"):
        self.config = self.load_config(config_path)
        
    def load_config(self, config_path):
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                return yaml.safe_load(f)
        except:
            return {}
    
    def parse_lot_numbers(self, lot_string):
        """
        تحليل نص أرقام اللوت واستخراج كل الأرقام
        تدعم: 139385, 139912/139913, 139912-139913, 139865/2, 139865/3, SFP228, 163-31-03-39-2394, DH956-TX/2025, 91191
        """
        lot_string = lot_string.strip()
        
        # النمط 1: رقمين مفصولين بـ / (مثال: 139912/139913)
        if '/' in lot_string and lot_string.replace('/', '').replace('-', '').isdigit():
            parts = lot_string.split('/')
            if len(parts) == 2 and parts[0].isdigit() and parts[1].isdigit():
                # لو الطرفين أرقام، يبقى دول لوطين منفصلين
                return {
                    "type": "explicit_multi",
                    "lots": [parts[0], parts[1]],
                    "count": 2,
                    "annotation_hint": None
                }
        
        # النمط 2: رقمين مفصولين بـ - (مثال: 139912-139913)
        if '-' in lot_string:
            # نتحقق لو ده رقم واحد طويل ولا رقمين
            parts = lot_string.split('-')
            # لو كل الأجزاء أرقام وطول كل جزء أكتر من 3 أرقام، يبقى دول رقمين
            if all(p.isdigit() and len(p) >= 3 for p in parts):
                return {
                    "type": "explicit_multi",
                    "lots": parts,
                    "count": len(parts),
                    "annotation_hint": None
                }
        
        # النمط 3: رقم/عدد (مثال: 139865/2 أو 139865/3 أو 139865/5)
        if '/' in lot_string:
            parts = lot_string.split('/')
            if len(parts) == 2 and parts[0].isdigit() and parts[1].isdigit():
                base_lot = parts[0]
                count = int(parts[1])
                return {
                    "type": "implicit_multi",
                    "base_lot": base_lot,
                    "lots": [base_lot],  # بنرجع الرقم الأساسي بس، والـ ERP هيتعامل مع الباقي
                    "count": count,
                    "annotation_hint": f"+{count-1}"
                }
        
        # النمط 4: رقم واحد (مثال: 139385, 91191, SFP228, 163-31-03-39-2394, DH956-TX/2025)
        return {
            "type": "single",
            "lots": [lot_string],
            "count": 1,
            "annotation_hint": None
        }
    
    def extract_lot_from_filename(self, filename):
        """استخراج أرقام اللوت من اسم الملف"""
        logger.info(f"Extracting lot from filename: {filename}")
        
        # إزالة الامتداد
        name_without_ext = os.path.splitext(filename)[0]
        
        # البحث عن نمط "Lot Number : XXX" في اسم الملف
        # مثال: "Lot Number 139385 Basil.pdf" أو "Basil Lot Number 139385.pdf"
        lot_match = re.search(r'Lot\s*Number\s*[:_\s]*\s*([A-Za-z0-9\-\/]+)', name_without_ext, re.IGNORECASE)
        
        if lot_match:
            lot_string = lot_match.group(1).strip()
            logger.info(f"Found lot string: {lot_string}")
            
            parsed = self.parse_lot_numbers(lot_string)
            logger.info(f"✓✓✓ PARSED LOTS: {parsed['lots']} (type: {parsed['type']}) ✓✓✓")
            return {
                "lot_raw": lot_string,
                "lot_parsed": parsed
            }
        
        # لو ملقناش "Lot Number"، ندور على أي رقم كبير (5-7 أرقام) أو نمط مختلط
        # نمط: أرقام-أرقام-أرقام (مثل 163-31-03-39-2394)
        complex_match = re.search(r'(\d{2,6}(?:-\d{2,6})+)', name_without_ext)
        if complex_match:
            lot_string = complex_match.group(1)
            logger.info(f"Found complex lot: {lot_string}")
            return {
                "lot_raw": lot_string,
                "lot_parsed": {
                    "type": "single",
                    "lots": [lot_string],
                    "count": 1,
                    "annotation_hint": None
                }
            }
        
        # نمط: أحرف وأرقام (مثل SFP228, DH956-TX/2025)
        alphanumeric_match = re.search(r'([A-Z]{2,}\d+[A-Z0-9\-\/]*)', name_without_ext, re.IGNORECASE)
        if alphanumeric_match:
            lot_string = alphanumeric_match.group(1)
            logger.info(f"Found alphanumeric lot: {lot_string}")
            return {
                "lot_raw": lot_string,
                "lot_parsed": {
                    "type": "single",
                    "lots": [lot_string],
                    "count": 1,
                    "annotation_hint": None
                }
            }
        
        # أخيراً: أي رقم 5-7 أرقام
        number_match = re.search(r'(\d{5,7})', name_without_ext)
        if number_match:
            lot_string = number_match.group(1)
            logger.info(f"Found numeric lot: {lot_string}")
            return {
                "lot_raw": lot_string,
                "lot_parsed": {
                    "type": "single",
                    "lots": [lot_string],
                    "count": 1,
                    "annotation_hint": None
                }
            }
        
        logger.error(f"✗✗✗ NO LOT NUMBER FOUND IN FILENAME: {filename} ✗✗✗")
        return None
    
    def extract_product_name(self, filename):
        """استخراج اسم المنتج من اسم الملف"""
        name_without_ext = os.path.splitext(filename)[0]
        
        # قائمة المنتجات المعروفة
        products = ['Basil', 'Fennel', 'Peppermint', 'Marjoram', 'Sage', 'Thyme', 
                   'Rosemary', 'Oregano', 'Parsley', 'Cilantro', 'Dill', 'Chamomile',
                   'Hibiscus', 'Calendula', 'Lavender', 'Melissa']
        
        for product in products:
            if product.lower() in name_without_ext.lower():
                return product
        
        # لو ملقناش، خد أول كلمة قبل "Lot"
        match = re.search(r'^([A-Za-z]+)', name_without_ext)
        if match:
            return match.group(1)
        
        return "UNKNOWN"
    
    def process_certificate(self, cert_path):
        logger.info(f"Processing: {os.path.basename(cert_path)}")
        
        filename = os.path.basename(cert_path)
        
        # استخراج اللوت من الاسم
        lot_data = self.extract_lot_from_filename(filename)
        
        if not lot_data:
            logger.error("FAILED - No lot found in filename")
            return None
        
        product_name = self.extract_product_name(filename)
        parsed = lot_data["lot_parsed"]
        
        # إعداد lot_info للـ ERP
        lot_info_list = []
        for lot in parsed["lots"]:
            lot_info_list.append({
                "num": lot,
                "type": parsed["type"],
                "base_lot": parsed.get("base_lot"),
                "count": parsed.get("count", 1),
                "annotation_hint": parsed.get("annotation_hint")
            })
        
        result = {
            "file_path": cert_path,
            "file_name": filename,
            "certification_number": "UNKNOWN",
            "product_name": product_name,
            "lot_numbers": parsed["lots"],  # قائمة بكل الأرقام اللي هندور عليها
            "lot_info": lot_info_list,      # معلومات إضافية لكل رقم
            "lot_structure": parsed["type"],
            "total_count": parsed["count"],
            "annotation_hint": parsed.get("annotation_hint"),
            "extraction_time": datetime.now().isoformat(),
        }
        
        logger.info(f"SUCCESS: Lots={result['lot_numbers']}, Structure={parsed['type']}, Product={product_name}")
        return result
    
    def run(self):
        logger.info("=== ExtractLotAgent (Filename-Based - Multi-Lot Support) ===")
        
        cert_inbox = self.config.get('paths', {}).get('cert_inbox', 'InPut/Cert_Inbox')
        
        if not os.path.exists(cert_inbox):
            logger.error(f"Inbox not found: {cert_inbox}")
            return []
        
        cert_files = [f for f in os.listdir(cert_inbox) if f.lower().endswith('.pdf')]
        logger.info(f"Found {len(cert_files)} PDF(s)")
        
        results = []
        for filename in cert_files:
            cert_path = os.path.join(cert_inbox, filename)
            result = self.process_certificate(cert_path)
            if result:
                results.append(result)
        
        logger.info(f"=== COMPLETED: {len(results)} successful ===")
        return results


def extract_lots_from_certificates(config_path="config.yaml"):
    agent = ExtractLotAgent(config_path)
    return agent.run()