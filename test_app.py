#!/usr/bin/env python3
"""
Simple test script to verify the Flask application functionality
"""

import os
import sys
from pathlib import Path

def test_imports():
    """Test if all required modules can be imported"""
    try:
        import flask
        import pandas
        import openpyxl
        from docx import Document
        from bs4 import BeautifulSoup
        print("✅ All required modules imported successfully")
        return True
    except ImportError as e:
        print(f"❌ Import error: {e}")
        return False

def test_extractors():
    """Test if our extractor modules work"""
    try:
        from document_extractor import extract_itac_report, extract_general_info_fields, extract_energy_usage
        from excel_keyinfo_extractor import extract_all_structured_info
        print("✅ Extractor modules imported successfully")
        return True
    except ImportError as e:
        print(f"❌ Extractor import error: {e}")
        return False

def test_app_creation():
    """Test if Flask app can be created"""
    try:
        from app import app
        print("✅ Flask app created successfully")
        print(f"   App name: {app.name}")
        print(f"   Debug mode: {app.debug}")
        return True
    except Exception as e:
        print(f"❌ Flask app creation error: {e}")
        return False

def test_sample_files():
    """Check if sample files exist"""
    docx_path = Path("docs/report1/LS2502 - Final Draft R2.docx")
    excel_path = Path("docs/report1/IACAssessmentTemplate.v2.1.xlsx")
    
    docx_exists = docx_path.exists()
    excel_exists = excel_path.exists()
    
    print(f"{'✅' if docx_exists else '❌'} DOCX sample file: {docx_path}")
    print(f"{'✅' if excel_exists else '❌'} Excel sample file: {excel_path}")
    
    return docx_exists and excel_exists

def test_extraction():
    """Test actual data extraction if sample files exist"""
    docx_path = Path("docs/report1/LS2502 - Final Draft R2.docx")
    excel_path = Path("docs/report1/IACAssessmentTemplate.v2.1.xlsx")
    
    if not (docx_path.exists() and excel_path.exists()):
        print("⚠️  Sample files not found, skipping extraction test")
        return True
    
    try:
        from document_extractor import extract_itac_report, extract_general_info_fields
        from excel_keyinfo_extractor import extract_all_structured_info
        
        # Test DOCX extraction
        print("Testing DOCX extraction...")
        doc_data = extract_itac_report(str(docx_path), output="html", save_files=False)
        doc_general = extract_general_info_fields(doc_data["general_information"])
        print(f"   Extracted {len(doc_general)} general info fields from DOCX")
        
        # Test Excel extraction
        print("Testing Excel extraction...")
        excel_data = extract_all_structured_info(str(excel_path))
        excel_general = excel_data.get("general_info", {})
        print(f"   Extracted {len(excel_general)} general info fields from Excel")
        
        print("✅ Data extraction test successful")
        return True
        
    except Exception as e:
        print(f"❌ Extraction test error: {e}")
        return False

def test_comparison():
    """Test comparison logic"""
    try:
        from app import compare_values, compare_general_info
        
        # Test numeric comparison
        result1 = compare_values(100.0, 99.5, tolerance=0.01)
        result2 = compare_values("Test String", "test string")
        result3 = compare_values(100, None)
        
        print("✅ Comparison logic test successful")
        print(f"   Numeric comparison (within tolerance): {result1['match']}")
        print(f"   String comparison (case insensitive): {result2['match']}")
        print(f"   Missing value comparison: {result3['mismatch_type']}")
        
        return True
        
    except Exception as e:
        print(f"❌ Comparison test error: {e}")
        return False

def main():
    """Run all tests"""
    print("🧪 ITAC Report Validator - Test Suite")
    print("=" * 50)
    
    tests = [
        ("Import Test", test_imports),
        ("Extractor Import Test", test_extractors),
        ("Flask App Creation Test", test_app_creation),
        ("Sample Files Check", test_sample_files),
        ("Data Extraction Test", test_extraction),
        ("Comparison Logic Test", test_comparison)
    ]
    
    passed = 0
    total = len(tests)
    
    for test_name, test_func in tests:
        print(f"\n🔍 Running {test_name}...")
        if test_func():
            passed += 1
        
    print("\n" + "=" * 50)
    print(f"📊 Test Results: {passed}/{total} tests passed")
    
    if passed == total:
        print("🎉 All tests passed! The application should work correctly.")
        print("\n🚀 To start the application, run:")
        print("   python app.py")
        print("\n🌐 Then open your browser to: http://localhost:5000")
    else:
        print("❌ Some tests failed. Please check the error messages above.")
        sys.exit(1)

if __name__ == "__main__":
    main()
