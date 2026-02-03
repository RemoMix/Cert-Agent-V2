import os
import time
import shutil
import tempfile
from datetime import datetime
import yaml
import fitz  # PyMuPDF
from PIL import Image, ImageDraw, ImageFont
import logging

# For Windows printing
try:
    import win32print
    import win32api
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False
    logging.warning("win32print not available")

logger = logging.getLogger('CertPrintAgent')

class AnnotatePrintAgent:
    def __init__(self, config_path="config.yaml"):
        self.config = self.load_config(config_path)
        self.printer_name = self.config.get('printing', {}).get('printer_name', '')
        self.retry_attempts = self.config.get('printing', {}).get('retry_attempts', 3)
        self.retry_delay = self.config.get('printing', {}).get('retry_delay_seconds', 10)
        self.annotation_config = self.config.get('printing', {}).get('annotation', {
            'position_x': 50,
            'position_y': 750,
            'font_size': 11
        })
        
        # Setup paths
        self.setup_paths()
        
    def load_config(self, config_path):
        """Load configuration file"""
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                return yaml.safe_load(f)
        except Exception as e:
            logger.error(f"Error loading config: {e}")
            return {}
    
    def setup_paths(self):
        """Setup all required directories"""
        base_dir = self.config.get('paths', {}).get('base_dir', 'Cert-Print-Agent')
        paths_config = self.config.get('paths', {})
        
        # New structure paths
        self.source_cert_dir = os.path.join(base_dir, paths_config.get('source_cert', 'GetCertAgent/Source_Cert'))
        self.annotated_dir = os.path.join(base_dir, paths_config.get('annotated_cert', 'GetCertAgent/Annotated_Certificates'))
        self.printed_dir = os.path.join(base_dir, paths_config.get('printed_cert', 'GetCertAgent/Printed_Annotated_Cert'))
        self.processed_dir = os.path.join(base_dir, paths_config.get('processed', 'GetCertAgent/Processed'))
        
        # Create directories
        for d in [self.source_cert_dir, self.annotated_dir, self.printed_dir, self.processed_dir]:
            os.makedirs(d, exist_ok=True)
            logger.debug(f"Directory ready: {d}")
    
    def is_printer_available(self):
        """Check if printer is available"""
        if not WIN32_AVAILABLE:
            return False
            
        try:
            printers = [printer[2] for printer in win32print.EnumPrinters(2)]
            
            if not self.printer_name:
                logger.warning("No printer configured")
                return False
            
            if self.printer_name in printers:
                logger.info(f"Printer available: {self.printer_name}")
                return True
            else:
                default_printer = win32print.GetDefaultPrinter()
                if default_printer:
                    logger.info(f"Using default printer: {default_printer}")
                    self.printer_name = default_printer
                    return True
                
                logger.warning(f"Printer not found: {self.printer_name}")
                return False
                
        except Exception as e:
            logger.error(f"Error checking printer: {e}")
            return False
    
    def annotate_pdf(self, pdf_path, annotation_text, output_dir=None):
        """Add annotation to PDF and save to annotated folder"""
        try:
            logger.info(f"Annotating: {os.path.basename(pdf_path)}")
            logger.info(f"Text: {annotation_text}")
            
            # Determine output path
            if output_dir is None:
                output_dir = self.annotated_dir
            
            filename = os.path.basename(pdf_path)
            base_name, ext = os.path.splitext(filename)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = f"{base_name}_{timestamp}_annotated{ext}"
            output_path = os.path.join(output_dir, output_filename)
            
            # Open PDF
            doc = fitz.open(pdf_path)
            page = doc[0]
            
            page_width = page.rect.width
            page_height = page.rect.height
            
            # Calculate position (bottom of page with margin)
            margin = 50
            y_position = page_height - margin
            
            # Create text box
            rect = fitz.Rect(
                self.annotation_config.get('position_x', margin),
                y_position - 25,
                page_width - margin,
                y_position + 5
            )
            
            # Draw white background
            page.draw_rect(rect, color=(1, 1, 1), fill=(1, 1, 1))
            
            # Add text with fallback methods
            try:
                page.insert_textbox(
                    rect,
                    annotation_text,
                    fontsize=self.annotation_config.get('font_size', 11),
                    color=(0, 0, 0),
                    align=0
                )
            except Exception as e:
                logger.warning(f"Textbox failed, trying freetext: {e}")
                annot = page.add_freetext_annot(
                    rect,
                    annotation_text,
                    fontsize=self.annotation_config.get('font_size', 11),
                    text_color=(0, 0, 0),
                    fill_color=(1, 1, 1),
                    border_color=(0, 0, 0)
                )
                annot.update()
            
            # Save annotated PDF
            doc.save(output_path)
            doc.close()
            
            logger.info(f"✓ Annotated PDF saved: {output_path}")
            return output_path
            
        except Exception as e:
            logger.error(f"Error annotating PDF: {e}")
            return None
    
    def print_pdf(self, pdf_path, retry=0):
        """Print PDF file"""
        if not WIN32_AVAILABLE:
            logger.error("Printing not available")
            return False
        
        try:
            logger.info(f"Printing (attempt {retry + 1}): {os.path.basename(pdf_path)}")
            
            result = win32api.ShellExecute(
                0,
                "print",
                pdf_path,
                f'/d:"{self.printer_name}"',
                ".",
                0
            )
            
            if result > 32:
                logger.info("✓ Print job sent")
                time.sleep(2)
                return True
            else:
                logger.error(f"Print failed with code: {result}")
                return False
                
        except Exception as e:
            logger.error(f"Error printing: {e}")
            return False
    
    def print_with_retry(self, pdf_path):
        """Print with retry logic"""
        for attempt in range(self.retry_attempts):
            if self.print_pdf(pdf_path, attempt):
                return True
            
            if attempt < self.retry_attempts - 1:
                logger.info(f"Waiting {self.retry_delay}s before retry...")
                time.sleep(self.retry_delay)
        
        logger.error(f"Failed after {self.retry_attempts} attempts")
        return False
    
    def organize_files(self, original_path, annotated_path, printed_success):
        """Organize files into appropriate folders"""
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = os.path.basename(original_path)
            base_name, ext = os.path.splitext(filename)
            
            # 1. Move original to Source_Cert
            source_filename = f"{base_name}_{timestamp}{ext}"
            source_path = os.path.join(self.source_cert_dir, source_filename)
            shutil.move(original_path, source_path)
            logger.info(f"✓ Original moved to Source_Cert: {source_filename}")
            
            # 2. If printed successfully, move annotated to Printed_Annotated_Cert
            if printed_success and annotated_path and os.path.exists(annotated_path):
                printed_filename = f"{base_name}_{timestamp}_printed{ext}"
                printed_dest = os.path.join(self.printed_dir, printed_filename)
                shutil.move(annotated_path, printed_dest)
                logger.info(f"✓ Printed copy saved: {printed_filename}")
                return printed_dest
            
            # 3. If not printed, keep in Annotated_Certificates
            elif annotated_path and os.path.exists(annotated_path):
                logger.info(f"✓ Annotated copy available: {os.path.basename(annotated_path)}")
                return annotated_path
            
            return None
            
        except Exception as e:
            logger.error(f"Error organizing files: {e}")
            return None
    
    def process_certificate(self, erp_result, original_pdf_path):
        """Process single certificate"""
        try:
            cert_number = erp_result.get('cert_number', 'UNKNOWN')
            annotation_text = erp_result.get('annotation_text', '')
            
            logger.info(f"\\nProcessing: {cert_number}")
            
            if not os.path.exists(original_pdf_path):
                logger.error(f"File not found: {original_pdf_path}")
                return False
            
            # Step 1: Annotate
            annotated_path = self.annotate_pdf(original_pdf_path, annotation_text)
            if not annotated_path:
                logger.error("Annotation failed")
                return False
            
            # Step 2: Print if available
            print_success = False
            if self.is_printer_available():
                print_success = self.print_with_retry(annotated_path)
            else:
                logger.warning("Printer not available, saving annotated only")
            
            # Step 3: Organize files
            final_path = self.organize_files(original_pdf_path, annotated_path, print_success)
            
            result_status = "printed" if print_success else "annotated"
            logger.info(f"✓ Completed: {result_status.upper()}")
            
            return print_success
            
        except Exception as e:
            logger.error(f"Error processing certificate: {e}")
            return False
    
    def process_all(self, erp_results):
        """Process all ERP results"""
        logger.info(f"\\nProcessing {len(erp_results)} certificate(s)")
        
        results = {
            'total': len(erp_results),
            'printed': 0,
            'annotated_only': 0,
            'failed': 0
        }
        
        for erp_result in erp_results:
            original_path = erp_result.get('file_path', '')
            
            success = self.process_certificate(erp_result, original_path)
            
            if success:
                results['printed'] += 1
            elif os.path.exists(erp_result.get('file_path', '')) == False:
                # File was processed but maybe not printed
                results['annotated_only'] += 1
            else:
                results['failed'] += 1
        
        logger.info(f"\\nSummary:")
        logger.info(f"  Total: {results['total']}")
        logger.info(f"  Printed: {results['printed']}")
        logger.info(f"  Annotated only: {results['annotated_only']}")
        logger.info(f"  Failed: {results['failed']}")
        
        return results
    
    def run(self, erp_results=None):
        """Run the agent"""
        logger.info("Starting AnnotatePrintAgent...")
        
        if not erp_results:
            logger.info("No results to process")
            return None
        
        return self.process_all(erp_results)

def annotate_and_print(erp_results, config_path="config.yaml"):
    agent = AnnotatePrintAgent(config_path)
    return agent.run(erp_results)
