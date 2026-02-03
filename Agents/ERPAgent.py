
# Update ERPAgent to generate proper annotation format
import pandas as pd
import os
import yaml
from datetime import datetime
import logging

logger = logging.getLogger('CertPrintAgent')

class ERPAgent:
    def __init__(self, config_path="config.yaml"):
        self.config = self.load_config(config_path)
        self.excel_path = self.get_excel_path()
        self.sheets = self.config.get('excel', {}).get('sheets', ["2026", "2025", "2024", "2023"])
        self.column_names = self.config.get('excel', {}).get('columns', {
            'cert_lot': 'NO',
            'internal_lot': 'Lot Num.',
            'supplier': 'Supplier'
        })
        self.excel_cache = {}
        
    def load_config(self, config_path):
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                return yaml.safe_load(f)
        except Exception as e:
            logger.error(f"Error loading config: {e}")
            return {}
    
    def get_excel_path(self):
        base_dir = self.config.get('paths', {}).get('base_dir', 'Cert-Print-Agent')
        erp_file = self.config.get('paths', {}).get('erp_file', 'Data/Raw_Warehouses.xlsx')
        return os.path.join(base_dir, erp_file)
    
    def load_excel_sheet(self, sheet_name):
        try:
            if sheet_name in self.excel_cache:
                return self.excel_cache[sheet_name]
            
            logger.info(f"Loading sheet: {sheet_name}")
            df = pd.read_excel(self.excel_path, sheet_name=sheet_name)
            df.columns = df.columns.str.strip()
            
            required = [
                self.column_names.get('cert_lot', 'NO'),
                self.column_names.get('internal_lot', 'Lot Num.'),
                self.column_names.get('supplier', 'Supplier')
            ]
            
            missing = [col for col in required if col not in df.columns]
            if missing:
                logger.warning(f"Missing columns in {sheet_name}: {missing}")
                return None
            
            cert_col = self.column_names.get('cert_lot', 'NO')
            df[cert_col] = df[cert_col].astype(str).str.strip().str.replace('.0', '', regex=False)
            
            self.excel_cache[sheet_name] = df
            logger.info(f"Sheet {sheet_name} loaded: {len(df)} rows")
            return df
            
        except Exception as e:
            logger.error(f"Error loading sheet {sheet_name}: {e}")
            return None
    
    def search_lot_in_sheet(self, cert_lot_number, sheet_name):
        try:
            df = self.load_excel_sheet(sheet_name)
            if df is None:
                return None
            
            cert_col = self.column_names.get('cert_lot', 'NO')
            lot_str = str(cert_lot_number).strip()
            
            matches = df[df[cert_col] == lot_str]
            
            if not matches.empty:
                return matches.iloc[0]
            
            # Try numeric
            try:
                lot_num = float(lot_str)
                matches = df[df[cert_col].astype(str).str.replace('.0', '', regex=False) == lot_str]
                if not matches.empty:
                    return matches.iloc[0]
            except:
                pass
            
            return None
            
        except Exception as e:
            logger.error(f"Error searching in {sheet_name}: {e}")
            return None
    
    def search_lot(self, cert_lot_number):
        logger.info(f"Searching ERP for lot: {cert_lot_number}")
        
        result = {
            'cert_lot': cert_lot_number,
            'found': False,
            'supplier': None,
            'internal_lot': None,
            'sheet_found': None
        }
        
        for sheet in self.sheets:
            row = self.search_lot_in_sheet(cert_lot_number, sheet)
            if row is not None:
                result['found'] = True
                result['supplier'] = str(row[self.column_names.get('supplier', 'Supplier')])
                result['internal_lot'] = str(row[self.column_names.get('internal_lot', 'Lot Num.')])
                result['sheet_found'] = sheet
                logger.info(f"Found in {sheet}: {result['supplier']} - {result['internal_lot']}")
                return result
        
        logger.warning(f"Lot {cert_lot_number} not found in ERP")
        return result
    
    def process_lot_with_context(self, lot_num, lot_info):
        info = lot_info if isinstance(lot_info, dict) else {'type': 'single', 'num': lot_num}
        
        if info.get('type') == 'implicit':
            base_result = self.search_lot(lot_num)
            if base_result['found']:
                base_result['implicit_count'] = info.get('count', 1)
                base_result['is_implicit'] = True
                return base_result
        
        return self.search_lot(lot_num)
    
    def search_multiple_lots(self, extraction_result):
        lot_numbers = extraction_result.get('lot_numbers', [])
        lot_info_list = extraction_result.get('lot_info', [])
        
        if not lot_numbers:
            return []
        
        logger.info(f"Searching for {len(lot_numbers)} lot(s)")
        results = []
        
        for i, lot_num in enumerate(lot_numbers):
            info = lot_info_list[i] if i < len(lot_info_list) else {'type': 'single', 'num': lot_num}
            result = self.process_lot_with_context(lot_num, info)
            results.append(result)
        
        return results
    
    def generate_annotation_text(self, lot_results):
        if not lot_results or not any(r.get('found') for r in lot_results):
            return "غير مسجل في النظام"
        
        parts = []
        for result in lot_results:
            if result.get('found'):
                supplier = result.get('supplier', '').strip()
                internal_lot = result.get('internal_lot', '').strip()
                
                if supplier and internal_lot:
                    # Format: "اسم المورد lot XXXX"
                    parts.append(f"{supplier}  {internal_lot}")
        
        if not parts:
            return "غير مسجل في النظام"
        
        return "\\n".join(parts)
    
    def process_certificate(self, extraction_result):
        cert_number = extraction_result.get('certification_number', 'UNKNOWN')
        logger.info(f"Processing cert: {cert_number}")
        
        lot_results = self.search_multiple_lots(extraction_result)
        
        found_count = sum(1 for r in lot_results if r.get('found'))
        all_found = found_count == len(lot_results) and len(lot_results) > 0
        
        annotation_text = self.generate_annotation_text(lot_results)
        
        result = {
            'cert_number': cert_number,
            'file_path': extraction_result.get('file_path', ''),
            'file_name': extraction_result.get('file_name', ''),
            'product': extraction_result.get('product_name', 'UNKNOWN'),
            'lot_results': lot_results,
            'annotation_text': annotation_text,
            'all_found': all_found,
            'processing_time': datetime.now().isoformat()
        }
        
        logger.info(f"ERP complete: {found_count}/{len(lot_results)} found")
        logger.info(f"Annotation: {annotation_text}")
        
        return result
    
    def process_all(self, extraction_results):
        logger.info(f"Processing {len(extraction_results)} certificates")
        return [self.process_certificate(ext) for ext in extraction_results]
    
    def run(self, extraction_results=None):
        logger.info("Starting ERPAgent...")
        if extraction_results:
            return self.process_all(extraction_results)
        return []

def process_erp_data(extraction_results, config_path="config.yaml"):
    agent = ERPAgent(config_path)
    return agent.run(extraction_results)

