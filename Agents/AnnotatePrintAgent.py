import os
import io
import time
import shutil
from datetime import datetime
import yaml
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import logging

try:
    import win32print
    import win32api
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False

logger = logging.getLogger('CertPrintAgent')

# ==================================================
# FONT
# ==================================================
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

FONT_PATHS = [
    os.path.join(BASE_DIR, "fonts", "arial.ttf"),
    "C:/Windows/Fonts/arial.ttf",
    "C:/Windows/Fonts/tahoma.ttf",
    "C:/Windows/Fonts/times.ttf",
]

FONT_PATH = None
for fp in FONT_PATHS:
    if os.path.exists(fp):
        FONT_PATH = fp
        break

if FONT_PATH:
    pdfmetrics.registerFont(TTFont("ArabicFont", FONT_PATH))
    logger.info(f"Font: {FONT_PATH}")
else:
    logger.error("No font found!")


class AnnotatePrintAgent:
    def __init__(self, config_path="config.yaml"):
        self.config = self.load_config(config_path)
        self.printer_name = self.config.get('printing', {}).get('printer_name', '')
        self.retry_attempts = self.config.get('printing', {}).get('retry_attempts', 3)
        self.retry_delay = self.config.get('printing', {}).get('retry_delay_seconds', 10)
        self.setup_paths()
        
    def load_config(self, config_path):
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                return yaml.safe_load(f)
        except Exception as e:
            logger.error(f"Error loading config: {e}")
            return {}
    
    def setup_paths(self):
        base_dir = self.config.get('paths', {}).get('base_dir', '.')
        paths_config = self.config.get('paths', {})
        
        self.source_cert_dir = os.path.join(base_dir, paths_config.get('source_cert', 'GetCertAgent/Source_Cert'))
        self.annotated_dir = os.path.join(base_dir, paths_config.get('annotated_cert', 'GetCertAgent/Annotated_Certificates'))
        self.printed_dir = os.path.join(base_dir, paths_config.get('printed_cert', 'GetCertAgent/Printed_Annotated_Cert'))
        self.cert_inbox = os.path.join(base_dir, paths_config.get('cert_inbox', 'GetCertAgent/Cert_Inbox'))
        
        for d in [self.source_cert_dir, self.annotated_dir, self.printed_dir]:
            os.makedirs(d, exist_ok=True)
    
    def build_annotated_pdf(self, pdf_path, annotation_text):
        """Build annotated PDF - سطرين منفصلين"""
        try:
            # تقسيم النص
            parts = annotation_text.split(' lot ')
            if len(parts) == 2:
                line1 = parts[0].strip()  # باسم حنا
                line2 = 'lot ' + parts[1].strip()  # lot 2601
            else:
                line1 = annotation_text
                line2 = ""
            
            # قراءة PDF
            reader = PdfReader(pdf_path)
            writer = PdfWriter()
            
            # إنشاء overlay
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=A4)
            
            font = "ArabicFont"
            size = 12
            can.setFont(font, size)
            
            # حساب أعراض الأسطر
            width1 = pdfmetrics.stringWidth(line1, font, size) if line1 else 0
            width2 = pdfmetrics.stringWidth(line2, font, size) if line2 else 0
            max_width = max(width1, width2)
            
            # الموقع (أعلى يمين)
            x = 560
            y = 800  # أعلى شوية عشان سطرين
            
            # خلفية رمادية للمستطيل
            box_height = 45 if line2 else 25
            can.setFillColorRGB(0.9, 0.9, 0.9)
            can.rect(x - max_width - 12, y - 5, max_width + 12, box_height, fill=1, stroke=0)
            
            # كتابة السطر الأول (العربي)
            can.setFillColorRGB(0, 0, 0)
            if line1:
                can.drawRightString(x - 6, y + 12, line1)
            
            # كتابة السطر الثاني (lot 2601)
            if line2:
                can.drawRightString(x - 6, y - 5, line2)
            
            can.save()
            packet.seek(0)
            
            # دمج
            overlay = PdfReader(packet)
            page = reader.pages[0]
            page.merge_page(overlay.pages[0])
            writer.add_page(page)
            
            for i in range(1, len(reader.pages)):
                writer.add_page(reader.pages[i])
            
            # حفظ
            filename = os.path.basename(pdf_path)
            base_name, ext = os.path.splitext(filename)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            out_pdf = os.path.join(self.annotated_dir, f"{base_name}_{timestamp}_ANNOTATED{ext}")
            
            with open(out_pdf, "wb") as f:
                writer.write(f)
            
            logger.info(f"✓ Annotated: {out_pdf}")
            return out_pdf
            
        except Exception as e:
            logger.error(f"Error: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return None
    
    def is_printer_available(self):
        if not WIN32_AVAILABLE:
            return False
        try:
            printers = [printer[2] for printer in win32print.EnumPrinters(2)]
            if self.printer_name in printers:
                return True
            default = win32print.GetDefaultPrinter()
            if default:
                self.printer_name = default
                return True
            return False
        except:
            return False
    
    def print_pdf(self, pdf_path, retry=0):
        if not WIN32_AVAILABLE:
            return False
        try:
            result = win32api.ShellExecute(0, "print", pdf_path, f'/d:"{self.printer_name}"', ".", 0)
            return result > 32
        except:
            return False
    
    def print_with_retry(self, pdf_path):
        for attempt in range(self.retry_attempts):
            if self.print_pdf(pdf_path, attempt):
                return True
            time.sleep(self.retry_delay)
        return False
    
    def find_pdf_file(self, filename):
        for path in [self.cert_inbox, self.source_cert_dir, '.']:
            full = os.path.join(path, filename)
            if os.path.exists(full):
                return full
        return None
    
    def process_certificate(self, erp_result, original_pdf_path):
        try:
            cert_number = erp_result.get('cert_number', 'UNKNOWN')
            annotation_text = erp_result.get('annotation_text', '')
            
            logger.info(f"Processing: {cert_number}")
            
            pdf_path = self.find_pdf_file(os.path.basename(original_pdf_path))
            if not pdf_path:
                logger.error(f"PDF not found")
                return False
            
            annotated = self.build_annotated_pdf(pdf_path, annotation_text)
            if not annotated:
                return False
            
            printed = False
            if self.is_printer_available():
                printed = self.print_with_retry(annotated)
            
            # نقل الملفات
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            new_name = f"{os.path.splitext(os.path.basename(pdf_path))[0]}_{timestamp}.pdf"
            shutil.move(pdf_path, os.path.join(self.source_cert_dir, new_name))
            
            if printed:
                printed_name = f"{os.path.splitext(os.path.basename(pdf_path))[0]}_{timestamp}_printed.pdf"
                shutil.copy(annotated, os.path.join(self.printed_dir, printed_name))
            
            return printed
            
        except Exception as e:
            logger.error(f"Error: {e}")
            return False
    
    def process_all(self, erp_results):
        logger.info(f"Processing {len(erp_results)} certificates")
        results = {'total': len(erp_results), 'printed': 0, 'annotated': 0, 'failed': 0}
        
        for erp in erp_results:
            success = self.process_certificate(erp, erp.get('file_path', ''))
            if success:
                results['printed'] += 1
            else:
                results['annotated'] += 1
        
        return results
    
    def run(self, erp_results=None):
        if not erp_results:
            return None
        return self.process_all(erp_results)


def annotate_and_print(erp_results, config_path="config.yaml"):
    agent = AnnotatePrintAgent(config_path)
    return agent.run(erp_results)
