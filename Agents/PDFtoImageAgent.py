import os
import json
import pytesseract
from pdf2image import convert_from_path
from PIL import Image, ImageEnhance, ImageFilter
import yaml
from datetime import datetime
from pathlib import Path
import logging

logger = logging.getLogger('CertPrintAgent')


class PDFtoImageAgent:
    def __init__(self, config_path="config.yaml"):
        self.config = self.load_config(config_path)
        self.setup_tesseract()
        self.setup_poppler()
        
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
        except Exception as e:
            logger.error(f"Error setting up Tesseract: {e}")
    
    def setup_poppler(self):
        self.poppler_path = self.config.get('paths', {}).get('poppler_path')
        if self.poppler_path and os.path.exists(self.poppler_path):
            logger.info(f"Poppler set from: {self.poppler_path}")
        else:
            self.poppler_path = None
    
    def preprocess_image(self, image):
        """تحسين الصورة قبل OCR"""
        # تحويل للرمادي
        gray = image.convert('L')
        # زيادة التباين
        enhancer = ImageEnhance.Contrast(gray)
        enhanced = enhancer.enhance(2.0)
        return enhanced
    
    def pdf_to_image(self, pdf_path):
        """Convert PDF to image(s)"""
        try:
            logger.info(f"Converting PDF: {os.path.basename(pdf_path)}")
            kwargs = {'dpi': 300}
            if self.poppler_path:
                kwargs['poppler_path'] = self.poppler_path
            
            images = convert_from_path(pdf_path, **kwargs)
            logger.info(f"Created {len(images)} image(s)")
            return images
        except Exception as e:
            logger.error(f"Error converting PDF: {e}")
            return []
    
    def ocr_image(self, image):
        """Extract text from image using OCR - IMPROVED"""
        try:
            # Preprocessing
            processed = self.preprocess_image(image)
            
            # محاولة 1: Arabic + English (الأفضل للشهادات المختلطة)
            try:
                text = pytesseract.image_to_string(processed, lang='ara+eng')
                if text.strip() and len(text) > 100:
                    logger.info(f"OCR (ara+eng): {len(text)} chars")
                    return text
            except Exception as e:
                logger.warning(f"ara+eng failed: {e}")
            
            # محاولة 2: English فقط
            text = pytesseract.image_to_string(processed, lang='eng')
            logger.info(f"OCR (eng): {len(text)} chars")
            return text
            
        except Exception as e:
            logger.error(f"OCR error: {e}")
            return ""
    
    def process_pdf(self, pdf_path, output_dir="temp_images"):
        """Process single PDF: Image → OCR → JSON"""
        filename = os.path.basename(pdf_path)
        base_name = os.path.splitext(filename)[0]
        
        logger.info(f"\\n{'='*50}")
        logger.info(f"Processing: {filename}")
        logger.info(f"{'='*50}")
        
        # Create directories
        os.makedirs(output_dir, exist_ok=True)
        json_dir = Path(output_dir) / "json"
        os.makedirs(json_dir, exist_ok=True)
        
        # Convert to images
        images = self.pdf_to_image(pdf_path)
        if not images:
            logger.error("Failed to convert PDF")
            return None
        
        results = []
        for i, image in enumerate(images):
            # Save image
            image_filename = f"{base_name}_page{i+1}.png"
            image_path = os.path.join(output_dir, image_filename)
            image.save(image_path, 'PNG')
            logger.info(f"Image saved: {image_filename}")
            
            # OCR
            logger.info(f"Running OCR on page {i+1}...")
            text = self.ocr_image(image)
            
            result = {
                "pdf_file": filename,
                "page_number": i + 1,
                "image_file": image_filename,
                "image_path": image_path,
                "ocr_text": text,
                "processed_at": datetime.now().isoformat()
            }
            results.append(result)
            
            # DEBUG: اطبع جزء من النص
            if text:
                sample = text.replace("\\n", " ")[:200]
                logger.info(f"Text sample: {sample}...")
        
        # Save JSON
        json_filename = f"{base_name}_ocr.json"
        json_path = json_dir / json_filename
        
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(results, f, ensure_ascii=False, indent=2)
        
        logger.info(f"JSON saved: {json_path}")
        
        return {
            "pdf_file": filename,
            "pages": len(results),
            "json_file": str(json_path),
            "results": results
        }
    
    def process_all(self, cert_inbox=None):
        """Process all PDFs in inbox"""
        if cert_inbox is None:
            cert_inbox = self.config.get('paths', {}).get('cert_inbox', 'GetCertAgent/Cert_Inbox')
        
        if not os.path.exists(cert_inbox):
            logger.error(f"Inbox not found: {cert_inbox}")
            return []
        
        pdf_files = [f for f in os.listdir(cert_inbox) if f.lower().endswith('.pdf')]
        logger.info(f"\\nFound {len(pdf_files)} PDF(s) to process")
        
        all_results = []
        for pdf_file in pdf_files:
            pdf_path = os.path.join(cert_inbox, pdf_file)
            result = self.process_pdf(pdf_path)
            if result:
                all_results.append(result)
        
        logger.info(f"\\n{'='*50}")
        logger.info(f"Total processed: {len(all_results)} PDFs")
        logger.info(f"{'='*50}")
        
        return all_results
    
    def run(self):
        """Run the agent"""
        logger.info("=== PDF to Image + OCR Agent ===")
        return self.process_all()


def convert_pdfs_to_json(config_path="config.yaml"):
    agent = PDFtoImageAgent(config_path)
    return agent.run()

