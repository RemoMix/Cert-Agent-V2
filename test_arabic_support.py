#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Test script for Arabic text rendering
سكريبت اختبار للتحقق من طباعة النص العربي بشكل صحيح
"""

import sys
import os

def test_arabic_support():
    """Test if Arabic support libraries are installed"""
    print("=" * 60)
    print("Testing Arabic Support / اختبار دعم اللغة العربية")
    print("=" * 60)
    
    # Test 1: Check imports
    print("\n1. Checking required libraries...")
    try:
        from arabic_reshaper import reshape
        # Try new bidi location first (v0.6+), then fall back to old location
        try:
            from bidi.bidi import get_display
            print("   ✓ arabic-reshaper: INSTALLED")
            print("   ✓ python-bidi: INSTALLED (v0.6+)")
        except ImportError:
            from bidi.algorithm import get_display
            print("   ✓ arabic-reshaper: INSTALLED")
            print("   ✓ python-bidi: INSTALLED (older version)")
        arabic_available = True
    except ImportError as e:
        print("   ✗ Missing libraries!")
        print("   Error:", str(e))
        print("\n   Install with:")
        print("   pip install arabic-reshaper python-bidi")
        arabic_available = False
    
    # Test 2: Test text reshaping
    if arabic_available:
        print("\n2. Testing text reshaping...")
        
        test_names = [
            "عزمي ابراهيم",
            "باسم حنا",
            "محمد علي",
            "احمد محمود"
        ]
        
        for name in test_names:
            reshaped = reshape(name)
            bidi = get_display(reshaped)
            print(f"   Original: {name}")
            print(f"   Reshaped: {reshaped}")
            print(f"   Display:  {bidi}")
            print()
    
    # Test 3: Test reportlab
    print("\n3. Checking PDF library...")
    try:
        from reportlab.pdfgen import canvas
        from reportlab.pdfbase import pdfmetrics
        print("   ✓ reportlab: INSTALLED")
    except ImportError:
        print("   ✗ reportlab: NOT INSTALLED")
        print("   Install with: pip install reportlab")
    
    # Test 4: Test PyPDF2
    print("\n4. Checking PDF reader...")
    try:
        from PyPDF2 import PdfReader, PdfWriter
        print("   ✓ PyPDF2: INSTALLED")
    except ImportError:
        print("   ✗ PyPDF2: NOT INSTALLED")
        print("   Install with: pip install PyPDF2")
    
    print("\n" + "=" * 60)
    if arabic_available:
        print("✓ All tests passed! / جميع الاختبارات نجحت!")
        print("You can use the improved AnnotatePrintAgent.")
    else:
        print("✗ Some tests failed / بعض الاختبارات فشلت")
        print("Please install missing packages.")
    print("=" * 60)


def test_annotation_text():
    """Test the annotation text parsing"""
    print("\n" + "=" * 60)
    print("Testing Annotation Text Parsing")
    print("=" * 60)
    
    test_cases = [
        "عزمي ابراهيم lot 2601",
        "باسم حنا lot 1234",
        "محمد علي lot 5678",
    ]
    
    for annotation_text in test_cases:
        print(f"\nInput: {annotation_text}")
        
        # Parse like in the agent
        parts = annotation_text.split(' lot ')
        if len(parts) == 2:
            arabic_name = parts[0].strip()
            lot_number = parts[1].strip()
            lot_text = f'Lot {lot_number}'
            
            print(f"  Arabic Name: {arabic_name}")
            print(f"  Lot Text:    {lot_text}")
            
            # Test reshaping
            try:
                from arabic_reshaper import reshape
                try:
                    from bidi.bidi import get_display
                except ImportError:
                    from bidi.algorithm import get_display
                
                reshaped = reshape(arabic_name)
                bidi = get_display(reshaped)
                print(f"  Display:     {bidi}")
            except ImportError:
                print("  (Arabic reshaping not available)")


def create_test_pdf():
    """Create a simple test PDF with Arabic text"""
    print("\n" + "=" * 60)
    print("Creating Test PDF")
    print("=" * 60)
    
    try:
        from reportlab.pdfgen import canvas
        from reportlab.lib.pagesizes import A4
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        from arabic_reshaper import reshape
        try:
            from bidi.bidi import get_display
        except ImportError:
            from bidi.algorithm import get_display
        import os
        
        # Find font
        font_paths = [
            "C:/Windows/Fonts/arial.ttf",
            "C:/Windows/Fonts/tahoma.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        ]
        
        font_path = None
        for fp in font_paths:
            if os.path.exists(fp):
                font_path = fp
                break
        
        if not font_path:
            print("   ✗ No Arabic font found!")
            return
        
        # Register font
        pdfmetrics.registerFont(TTFont("ArabicFont", font_path))
        print(f"   Using font: {font_path}")
        
        # Create PDF
        output_file = "test_arabic_annotation.pdf"
        c = canvas.Canvas(output_file, pagesize=A4)
        c.setFont("ArabicFont", 14)
        
        # Test data
        arabic_name = "عزمي ابراهيم"
        lot_text = "Lot 2601"
        
        # Prepare Arabic text
        arabic_display = get_display(reshape(arabic_name))
        
        # Draw on PDF
        x = 560
        y1 = 815
        y2 = 790
        
        # Draw backgrounds
        c.setFillColorRGB(0.9, 0.9, 0.9)
        c.rect(x - 150, y1 - 3, 145, 20, fill=1, stroke=0)
        c.rect(x - 150, y2 - 3, 145, 20, fill=1, stroke=0)
        
        # Draw text
        c.setFillColorRGB(0, 0, 0)
        c.drawRightString(x - 5, y1, arabic_display)
        c.drawRightString(x - 5, y2, lot_text)
        
        c.save()
        print(f"   ✓ Test PDF created: {output_file}")
        print(f"   Open it to verify Arabic text rendering")
        
    except ImportError as e:
        print(f"   ✗ Missing library: {e}")
    except Exception as e:
        print(f"   ✗ Error: {e}")


if __name__ == "__main__":
    test_arabic_support()
    test_annotation_text()
    
    # Ask if user wants to create test PDF
    print("\n" + "=" * 60)
    response = input("Create test PDF? (y/n): ")
    if response.lower() == 'y':
        create_test_pdf()
    
    print("\n✓ Testing complete!")
