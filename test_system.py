
#!/usr/bin/env python3
"""
Test script for Cert-Print-Agent
Run this to verify all components are working
"""

import os
import sys
import yaml
from datetime import datetime

# Add project root to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

def test_logging():
    """Test logging system"""
    print("\\n[TEST] Logging System")
    try:
        from Agents.LoggingAgent import get_logger
        logger = get_logger()
        logger.info("Test log message")
        print("✓ Logging working")
        return True
    except Exception as e:
        print(f"✗ Logging failed: {e}")
        return False

def test_ocr():
    """Test OCR system"""
    print("\\n[TEST] OCR System (Tesseract)")
    try:
        import pytesseract
        from PIL import Image
        
        # Create a simple test image
        img = Image.new('RGB', (100, 30), color='white')
        
        # Try OCR
        text = pytesseract.image_to_string(img)
        print(f"✓ OCR working (tesseract version: {pytesseract.get_tesseract_version()})")
        return True
    except Exception as e:
        print(f"✗ OCR failed: {e}")
        print("  Make sure Tesseract is installed and in PATH")
        return False

def test_pdf_processing():
    """Test PDF processing"""
    print("\\n[TEST] PDF Processing")
    try:
        from pdf2image import convert_from_path
        import fitz
        print("✓ PDF libraries imported successfully")
        return True
    except Exception as e:
        print(f"✗ PDF processing failed: {e}")
        return False

def test_excel():
    """Test Excel reading"""
    print("\\n[TEST] Excel Processing")
    try:
        import pandas as pd
        
        # Check if ERP file exists
        config_path = "config.yaml"
        if os.path.exists(config_path):
            with open(config_path, 'r', encoding='utf-8') as f:
                config = yaml.safe_load(f)
            
            erp_path = config.get('paths', {}).get('erp_file', 'Data/Raw_Warehouses.xlsx')
            base_dir = config.get('paths', {}).get('base_dir', 'Cert-Print-Agent')
            full_path = os.path.join(base_dir, erp_path)
            
            if os.path.exists(full_path):
                # Try reading first sheet
                df = pd.read_excel(full_path, sheet_name=0, nrows=5)
                print(f"✓ Excel file accessible ({len(df)} rows read)")
                print(f"  Columns: {list(df.columns[:5])}")
                return True
            else:
                print(f"⚠ ERP file not found at: {full_path}")
                return False
        else:
            print("⚠ Config file not found")
            return False
            
    except Exception as e:
        print(f"✗ Excel test failed: {e}")
        return False

def test_agents():
    """Test agent initialization"""
    print("\\n[TEST] Agent Initialization")
    results = []
    
    try:
        from Agents.ExtractLotAgent import ExtractLotAgent
        agent = ExtractLotAgent()
        print("✓ ExtractLotAgent initialized")
        results.append(True)
    except Exception as e:
        print(f"✗ ExtractLotAgent failed: {e}")
        results.append(False)
    
    try:
        from Agents.ERPAgent import ERPAgent
        agent = ERPAgent()
        print("✓ ERPAgent initialized")
        results.append(True)
    except Exception as e:
        print(f"✗ ERPAgent failed: {e}")
        results.append(False)
    
    try:
        from Agents.AnnotatePrintAgent import AnnotatePrintAgent
        agent = AnnotatePrintAgent()
        print("✓ AnnotatePrintAgent initialized")
        results.append(True)
    except Exception as e:
        print(f"✗ AnnotatePrintAgent failed: {e}")
        results.append(False)
    
    return all(results)

def test_directory_structure():
    """Test directory structure"""
    print("\\n[TEST] Directory Structure")
    
    required_dirs = [
        'logs',
        'GetCertAgent/MyEmails',
        'GetCertAgent/Cert_Inbox',
        'GetCertAgent/Processed',
        'GetCertAgent/TempImages',
        'Data'
    ]
    
    base_dir = 'Cert-Print-Agent'
    all_exist = True
    
    for d in required_dirs:
        path = os.path.join(base_dir, d)
        if os.path.exists(path):
            print(f"✓ {d}")
        else:
            print(f"✗ {d} (missing)")
            all_exist = False
            # Create it
            try:
                os.makedirs(path, exist_ok=True)
                print(f"  → Created: {path}")
            except Exception as e:
                print(f"  → Could not create: {e}")
    
    return all_exist

def test_config():
    """Test configuration file"""
    print("\\n[TEST] Configuration File")
    
    config_path = "config.yaml"
    if not os.path.exists(config_path):
        print(f"✗ Config file not found: {config_path}")
        return False
    
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config = yaml.safe_load(f)
        
        print("✓ Config file loaded")
        
        # Check required sections
        required = ['paths', 'excel', 'printing', 'ocr', 'monitoring', 'logging']
        missing = [r for r in required if r not in config]
        
        if missing:
            print(f"⚠ Missing sections: {missing}")
            return False
        
        print("  Sections: " + ", ".join(required))
        return True
        
    except Exception as e:
        print(f"✗ Config error: {e}")
        return False

def run_full_test():
    """Run all tests"""
    print("=" * 70)
    print("Cert-Print-Agent System Test")
    print("=" * 70)
    print(f"Time: {datetime.now()}")
    print(f"Python: {sys.version}")
    print(f"Platform: {sys.platform}")
    
    tests = [
        ("Configuration", test_config),
        ("Directory Structure", test_directory_structure),
        ("Logging System", test_logging),
        ("OCR (Tesseract)", test_ocr),
        ("PDF Processing", test_pdf_processing),
        ("Excel Processing", test_excel),
        ("Agents", test_agents),
    ]
    
    results = []
    for name, test_func in tests:
        try:
            result = test_func()
            results.append((name, result))
        except Exception as e:
            print(f"\\n✗ {name} crashed: {e}")
            results.append((name, False))
    
    # Summary
    print("\\n" + "=" * 70)
    print("TEST SUMMARY")
    print("=" * 70)
    
    passed = sum(1 for _, r in results if r)
    total = len(results)
    
    for name, result in results:
        status = "✓ PASS" if result else "✗ FAIL"
        print(f"{status:8} {name}")
    
    print("-" * 70)
    print(f"Total: {passed}/{total} tests passed")
    
    if passed == total:
        print("\\n🎉 All tests passed! System ready.")
        return 0
    else:
        print("\\n⚠ Some tests failed. Check errors above.")
        return 1

if __name__ == "__main__":
    exit(run_full_test())
'''

with open('/mnt/kimi/output/test_system.py', 'w', encoding='utf-8') as f:
    f.write(test_code)

print("✓ test_system.py تم إنشاؤه بنجاح")
