#!/usr/bin/env python3
"""
Comparison script to validate that excel_keyinfo_extractor and document_extractor
produce consistent results for the same IAC assessment data.
"""

from excel_keyinfo_extractor import extract_general_info_dict
from document_extractor import extract_itac_report, extract_general_info_fields
import json
from pathlib import Path

def compare_general_info():
    """Compare general info extraction between Excel and Document extractors."""
    
    # File paths
    excel_path = "docs/report1/IACAssessmentTemplate.v2.1.xlsx"
    docx_path = "docs/report1/LS2502 - Final Draft R2.docx"
    
    print("🔍 Comparing General Info extraction between Excel and Document sources...")
    
    # Extract from Excel
    excel_data = extract_general_info_dict(excel_path)
    
    # Extract from Document
    doc_data = extract_itac_report(docx_path, output="html", save_files=False)
    doc_general_info = extract_general_info_fields(doc_data["general_information"])
    
    print(f"\n📊 Extraction Results:")
    print(f"   Excel fields extracted: {len(excel_data)}")
    print(f"   Document fields extracted: {len(doc_general_info)}")
    
    # Find common fields
    excel_keys = set(excel_data.keys())
    doc_keys = set(doc_general_info.keys())
    common_keys = excel_keys.intersection(doc_keys)
    
    print(f"   Common standardized keys: {len(common_keys)}")
    
    # Compare values for common fields
    print(f"\n🔄 Field-by-field comparison:")
    matches = 0
    differences = 0
    
    for key in sorted(common_keys):
        excel_val = excel_data[key]
        doc_val = doc_general_info[key]
        
        # Handle numeric comparisons with tolerance
        if isinstance(excel_val, (int, float)) and isinstance(doc_val, (int, float)):
            match = abs(excel_val - doc_val) < 0.01  # Small tolerance for floating point
        else:
            match = str(excel_val).strip().lower() == str(doc_val).strip().lower()
        
        status = "✅" if match else "❌"
        print(f"   {status} {key}")
        print(f"      Excel: {excel_val}")
        print(f"      Doc:   {doc_val}")
        
        if match:
            matches += 1
        else:
            differences += 1
    
    # Show fields unique to each source
    excel_only = excel_keys - doc_keys
    doc_only = doc_keys - excel_keys
    
    if excel_only:
        print(f"\n📋 Fields only in Excel: {sorted(excel_only)}")
    
    if doc_only:
        print(f"\n📄 Fields only in Document: {sorted(doc_only)}")
    
    # Summary
    total_compared = matches + differences
    match_rate = (matches / total_compared * 100) if total_compared > 0 else 0
    
    print(f"\n📈 Summary:")
    print(f"   Fields compared: {total_compared}")
    print(f"   Matches: {matches}")
    print(f"   Differences: {differences}")
    print(f"   Match rate: {match_rate:.1f}%")
    
    # Save comparison results
    comparison_results = {
        "excel_data": excel_data,
        "document_data": doc_general_info,
        "comparison_summary": {
            "total_compared": total_compared,
            "matches": matches,
            "differences": differences,
            "match_rate_percent": match_rate,
            "excel_only_fields": list(excel_only),
            "document_only_fields": list(doc_only)
        }
    }
    
    with open("extractor_comparison_results.json", "w") as f:
        json.dump(comparison_results, f, indent=2, default=str)
    
    print(f"\n✅ Detailed comparison saved to: extractor_comparison_results.json")
    
    return comparison_results

if __name__ == "__main__":
    compare_general_info()
