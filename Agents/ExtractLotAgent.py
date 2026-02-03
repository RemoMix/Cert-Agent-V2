
# Fix ExtractLotAgent with better regex patterns
import os
import re
import pytesseract
from pdf2image import convert_from_path
from PIL import Image, ImageEnhance, ImageFilter
import yaml
from datetime import datetime
import logging

try:
    import cv2
    import numpy as np
    CV2_AVAILABLE = True
except ImportError:
    CV2_AVAILABLE = False

logger = logging.getLogger('CertPrintAgent')

class ExtractLotAgent:
    def __init__(self, config_path="config.yaml"):
        self.config = self.load_config(config_path)
        self.setup_tesseract()
        self.setup_poppler()
        self.lot_patterns = self.setup_lot_patterns()
        
    def load_config(self, config_path):
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                return yaml.safe_load(f)
        except Exception as e:
            logger.error(f"Error loading config: {e}")
            return {}
    
    def setup_tesseract(self):
        try:
            tesseract_path = self.config.get('paths', {}).get('tesseract_path', 'Tesseract/tesseract.exe')
            if os.path.exists(tesseract_path):
                pytesseract.pytesseract.tesseract_cmd = tesseract_path
                logger.info(f"Tesseract set from: {tesseract_path}")
            else:
                pytesseract.pytesseract.tesseract_cmd = 'tesseract'
        except Exception as e:
            logger.error(f"Error setting up Tesseract: {e}")
    
    def setup_poppler(self):
        try:
            poppler_path = self.config.get('paths', {}).get('poppler_path')
            if poppler_path and os.path.exists(poppler_path):
                self.poppler_path = poppler_path
                logger.info(f"Poppler set from: {poppler_path}")
            else:
                self.poppler_path = None
        except Exception as e:
            logger.error(f"Error setting up Poppler: {e}")
            self.poppler_path = None
    
    def setup_lot_patterns(self):
        """Enhanced lot number patterns - FIXED"""
        patterns = [
            # Pattern 1: "Lot Number : 139928" (most common)
            r'Lot\\s*Number\\s*[:：]?\\s*(\\d{5,7})\\b',
            
            # Pattern 2: "Lot: 139928" or "Lot:139928"
            r'Lot\\s*[:：]\\s*(\\d{5,7})\\b',
            
            # Pattern 3: "Lot No: 139928" or "Lot No.: 139928"
            r'Lot\\s*No\\.?\\s*[:：]?\\s*(\\d{5,7})\\b',
            
            # Pattern 4: "Lot# 139928"
            r'Lot\\s*#\\s*(\\d{5,7})\\b',
            
            # Pattern 5: Explicit multi-lot: 139912/139913
            r'Lot\\s*(?:Number|No\\.?)?\\s*[:：]?\\s*(\\d{5,7})\\s*[/\\-]\\s*(\\d{5,7})',
            
            # Pattern 6: Implicit multi-lot: 139865/2
            r'Lot\\s*(?:Number|No\\.?)?\\s*[:：]?\\s*(\\d{5,7})\\s*/\\s*(\\d{1,2})\\b',
            
            # Pattern 7: Arabic format
            r'رقم\\s*(?:اللوت|الباتش)\\s*[:：]?\\s*(\\d{5,7})',
            
            # Pattern 8: Alphanumeric lot
            r'Lot\\s*(?:Number|No\\.?)?\\s*[:：]?\\s*([A-Z]{2,3}\\d{3,4})\\b',
            
            # Pattern 9: Fallback - standalone 5-7 digit numbers
            r'\\b(\\d{5,7})\\b',
        ]
        return patterns
    
    def preprocess_image(self, image):
        try:
            if CV2_AVAILABLE:
                return self._preprocess_with_opencv(image)
            else:
                return self._preprocess_with_pil(image)
        except Exception as e:
            logger.warning(f"Image preprocessing failed: {e}, using original")
            return image
    
    def _preprocess_with_opencv(self, image):
        img_array = np.array(image)
        img_cv = cv2.cvtColor(img_array, cv2.COLOR_RGB2BGR)
        gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)
        denoised = cv2.fastNlMeansDenoising(gray, None, 10, 7, 21)
        clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))
        enhanced = clahe.apply(denoised)
        _, binary = cv2.threshold(enhanced, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        return Image.fromarray(binary)
    
    def _preprocess_with_pil(self, image):
        gray = image.convert('L')
        enhancer = ImageEnhance.Contrast(gray)
        enhanced = enhancer.enhance(2.0)
        sharpener = ImageEnhance.Sharpness(enhanced)
        sharpened = sharpener.enhance(2.0)
        binary = sharpened.point(lambda x: 0 if x < 128 else 255, '1')
        return binary.convert('L')
    
    def pdf_to_images(self, pdf_path):
        try:
            logger.info(f"Converting PDF to images: {os.path.basename(pdf_path)}")
            kwargs = {'dpi': 400}
            if self.poppler_path and os.path.exists(self.poppler_path):
                kwargs['poppler_path'] = self.poppler_path
            images = convert_from_path(pdf_path, **kwargs)
            logger.info(f"Converted PDF to {len(images)} images")
            return images
        except Exception as e:
            logger.error(f"Error converting PDF: {e}")
            return []
    
    def extract_text_from_image(self, image):
        try:
            processed_img = self.preprocess_image(image)
            languages = ['eng', 'eng+ara', 'ara']
            best_text = ""
            for lang in languages:
                try:
                    text = pytesseract.image_to_string(processed_img, lang=lang)
                    if len(text) > len(best_text):
                        best_text = text
                except:
                    continue
            if not best_text.strip():
                best_text = pytesseract.image_to_string(image, lang='eng')
            return best_text
        except Exception as e:
            logger.error(f"OCR error: {e}")
            return ""
    
    def extract_text_from_pdf(self, pdf_path):
        logger.info(f"Extracting text from: {os.path.basename(pdf_path)}")
        images = self.pdf_to_images(pdf_path)
        if not images:
            return ""
        all_text = ""
        for i, image in enumerate(images):
            logger.info(f"Processing page {i+1}/{len(images)}")
            text = self.extract_text_from_image(image)
            all_text += text + "\\n---PAGE BREAK---\\n"
        self.save_extracted_text(pdf_path, all_text)
        return all_text
    
    def save_extracted_text(self, pdf_path, text):
        try:
            debug_dir = os.path.join(self.config.get('paths', {}).get('base_dir', '.'), 'debug_texts')
            os.makedirs(debug_dir, exist_ok=True)
            base_name = os.path.basename(pdf_path)
            text_file = os.path.join(debug_dir, f"{base_name}.txt")
            with open(text_file, 'w', encoding='utf-8') as f:
                f.write(text)
            logger.info(f"Debug text saved: {text_file}")
        except:
            pass
    
    def extract_lot_numbers(self, text):
        """Extract lot numbers with improved logic"""
        lot_numbers = []
        lot_info = []
        
        logger.info("=== SEARCHING FOR LOT NUMBER ===")
        
        for pattern in self.lot_patterns:
            matches = list(re.finditer(pattern, text, re.IGNORECASE))
            if matches:
                logger.info(f"Pattern matched: {pattern[:50]}...")
            
            for match in matches:
                groups = match.groups()
                
                if len(groups) == 2:
                    first, second = groups[0], groups[1]
                    if len(second) >= 5:
                        # Explicit multi-lot
                        for num in [first, second]:
                            if num not in lot_numbers:
                                lot_numbers.append(num)
                                lot_info.append({'num': num, 'type': 'explicit_multi'})
                        logger.info(f"Found explicit multi-lot: {first}/{second}")
                    else:
                        # Implicit multi-lot
                        if first not in lot_numbers:
                            lot_numbers.append(first)
                            lot_info.append({'num': first, 'type': 'implicit', 'count': int(second)})
                        logger.info(f"Found implicit lot: {first}/{second}")
                
                elif len(groups) == 1:
                    lot_num = groups[0]
                    if lot_num not in lot_numbers:
                        # Validate it's not a false positive (like year 2026)
                        if not self._is_false_positive(lot_num, text):
                            lot_numbers.append(lot_num)
                            lot_info.append({'num': lot_num, 'type': 'single'})
                            logger.info(f"Found single lot: {lot_num}")
        
        return lot_numbers, lot_info
    
    def _is_false_positive(self, number, text):
        """Check if number is likely a false positive"""
        # Check if it's a year (2023, 2024, 2025, 2026)
        if number in ['2023', '2024', '2025', '2026']:
            return True
        # Check if it's preceded by keywords that suggest it's not a lot number
        patterns_to_exclude = [
            r'Date\\s*[:：]?\\s*' + number,
            r'Year\\s*[:：]?\\s*' + number,
            r'Page\\s*' + number,
            r'Certificate\\s*#?' + number,
        ]
        for pattern in patterns_to_exclude:
            if re.search(pattern, text, re.IGNORECASE):
                return True
        return False
    
    def extract_certification_number(self, text):
        patterns = [
            r'Certificate\\s*(?:Number|No\\.?)?\\s*[:：]\\s*([A-Za-z]+[-–]\\d+)',
            r'Cert\\.?\\s*#?\\s*[:：]?\\s*([A-Za-z0-9-]+)',
            r'(Dokki[-–]\\d+)',
            r'(ISM[-–]\\d+)',
        ]
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                cert_num = match.group(1).strip()
                if cert_num and len(cert_num) > 3:
                    logger.info(f"Found certification number: {cert_num}")
                    return cert_num
        return "UNKNOWN"
    
    def extract_product_name(self, text):
        # Try to get from "Sample : Basil" line
        sample_pattern = r'Sample\\s*[:：]\\s*([A-Za-z]{3,20})'
        match = re.search(sample_pattern, text, re.IGNORECASE)
        if match:
            product = match.group(1).strip()
            logger.info(f"Found product name: {product}")
            return product
        return "UNKNOWN"
    
    def process_certificate(self, cert_path):
        logger.info(f"Processing certificate: {os.path.basename(cert_path)}")
        
        text = self.extract_text_from_pdf(cert_path)
        if not text:
            logger.error(f"No text extracted from: {cert_path}")
            return None
        
        cert_number = self.extract_certification_number(text)
        product_name = self.extract_product_name(text)
        lot_numbers, lot_info = self.extract_lot_numbers(text)
        
        lot_type = self.determine_lot_type(lot_info)
        
        result = {
            "file_path": cert_path,
            "file_name": os.path.basename(cert_path),
            "certification_number": cert_number,
            "product_name": product_name,
            "lot_numbers": lot_numbers,
            "lot_info": lot_info,
            "lot_structure": lot_type,
            "extraction_time": datetime.now().isoformat(),
        }
        
        logger.info(f"Extraction complete:")
        logger.info(f"  - Cert: {cert_number}")
        logger.info(f"  - Product: {product_name}")
        logger.info(f"  - Lots: {lot_numbers}")
        logger.info(f"  - Structure: {lot_type}")
        
        return result
    
    def determine_lot_type(self, lot_info):
        if not lot_info:
            return "unknown"
        types = [info['type'] for info in lot_info]
        if 'implicit' in types:
            return "implicit_multi"
        elif types.count('explicit_multi') >= 2:
            return "explicit_multi"
        elif len(lot_info) == 1:
            return "single"
        else:
            return "multiple"
    
    def run(self):
        logger.info("Starting ExtractLotAgent...")
        cert_inbox = self.config.get('paths', {}).get('cert_inbox', 'GetCertAgent/Cert_Inbox')
        
        if not os.path.exists(cert_inbox):
            logger.error(f"Folder not found: {cert_inbox}")
            return []
        
        cert_files = []
        for ext in ['.pdf', '.jpg', '.jpeg', '.png', '.tiff', '.bmp']:
            cert_files.extend([
                os.path.join(cert_inbox, f)
                for f in os.listdir(cert_inbox)
                if f.lower().endswith(ext)
            ])
        
        if not cert_files:
            logger.info("No certificates found in Cert_Inbox")
            return []
        
        logger.info(f"Found {len(cert_files)} certificates to process")
        
        results = []
        for cert_path in cert_files:
            result = self.process_certificate(cert_path)
            if result:
                results.append(result)
        
        return results

def extract_lots_from_certificates(config_path="config.yaml"):
    agent = ExtractLotAgent(config_path)
    return agent.run()
