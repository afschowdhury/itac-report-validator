#!/usr/bin/env python3
"""
Enhanced comparison script to validate energy extraction between 
excel_keyinfo_extractor and document_extractor with standardized energy types.
"""

from excel_keyinfo_extractor import extract_all_structured_info
from document_extractor import extract_itac_report, extract_energy_usage
import json
from pathlib import Path

def compare_energy_data():
    """Compare energy data extraction between Excel and Document extractors."""
    
    # File paths
    excel_path = "docs/report1/IACAssessmentTemplate.v2.1.xlsx"
    docx_path = "docs/report1/LS2502 - Final Draft R2.docx"
    
    print("⚡ Comparing Energy Data extraction between Excel and Document sources...")
    
    # Extract from Excel (structured format)
    excel_data = extract_all_structured_info(excel_path)
    excel_energy = excel_data["energy_waste_info"]
    
    # Extract from Document
    doc_data = extract_itac_report(docx_path, output="html", save_files=False)
    doc_energy = extract_energy_usage(doc_data["annual_energy_usages_and_costs"])
    
    print(f"\n📊 Extraction Results:")
    print(f"   Excel energy entries: {len(excel_energy['data'])}")
    print(f"   Document energy entries: {len(doc_energy['data'])}")
    
    # Create lookup dictionaries by energy type
    excel_by_type = {item["type"]: item for item in excel_energy["data"]}
    doc_by_type = {item["type"]: item for item in doc_energy["data"]}
    
    excel_types = set(excel_by_type.keys())
    doc_types = set(doc_by_type.keys())
    common_types = excel_types.intersection(doc_types)
    
    print(f"   Common energy types: {len(common_types)}")
    
    # Compare energy types and values
    print(f"\n🔄 Energy Type Comparison:")
    cost_matches = 0
    cost_differences = 0
    usage_matches = 0
    usage_differences = 0
    
    for energy_type in sorted(common_types):
        excel_item = excel_by_type[energy_type]
        doc_item = doc_by_type[energy_type]
        
        print(f"\n   🔋 {energy_type}:")
        print(f"      Excel original: {excel_item.get('original_name', 'N/A')}")
        
        # Compare costs
        excel_cost = excel_item.get("cost", 0)
        doc_cost = doc_item.get("cost", 0)
        
        if isinstance(excel_cost, (int, float)) and isinstance(doc_cost, (int, float)):
            cost_match = abs(excel_cost - doc_cost) < 0.01
            cost_status = "✅" if cost_match else "❌"
            print(f"      {cost_status} Cost - Excel: ${excel_cost:,.2f}, Doc: ${doc_cost:,.2f}")
            
            if cost_match:
                cost_matches += 1
            else:
                cost_differences += 1
        else:
            print(f"      ⚠️ Cost - Excel: {excel_cost}, Doc: {doc_cost} (type mismatch)")
        
        # Compare usage data (more complex due to different structures)
        excel_usage = excel_item.get("usage", {})
        doc_usage = doc_item.get("usage", {})
        
        print(f"      📊 Usage - Excel: {excel_usage}, Doc: {doc_usage}")
        
        # Try to find matching usage values
        excel_usage_values = list(excel_usage.values()) if excel_usage else []
        doc_usage_values = list(doc_usage.values()) if doc_usage else []
        
        if excel_usage_values and doc_usage_values:
            # Compare first values (most likely to match)
            excel_val = excel_usage_values[0] if excel_usage_values else 0
            doc_val = doc_usage_values[0] if doc_usage_values else 0
            
            if isinstance(excel_val, (int, float)) and isinstance(doc_val, (int, float)):
                usage_match = abs(excel_val - doc_val) < 0.01
                usage_status = "✅" if usage_match else "❌"
                print(f"      {usage_status} Usage values similar")
                
                if usage_match:
                    usage_matches += 1
                else:
                    usage_differences += 1
        
    # Show unique energy types
    excel_only = excel_types - doc_types
    doc_only = doc_types - excel_types
    
    if excel_only:
        print(f"\n📋 Energy types only in Excel: {sorted(excel_only)}")
    
    if doc_only:
        print(f"\n📄 Energy types only in Document: {sorted(doc_only)}")
    
    # Summary statistics
    total_cost_compared = cost_matches + cost_differences
    cost_match_rate = (cost_matches / total_cost_compared * 100) if total_cost_compared > 0 else 0
    
    total_usage_compared = usage_matches + usage_differences  
    usage_match_rate = (usage_matches / total_usage_compared * 100) if total_usage_compared > 0 else 0
    
    print(f"\n📈 Summary:")
    print(f"   Energy types compared: {len(common_types)}")
    print(f"   Cost matches: {cost_matches}/{total_cost_compared} ({cost_match_rate:.1f}%)")
    print(f"   Usage matches: {usage_matches}/{total_usage_compared} ({usage_match_rate:.1f}%)")
    
    # Compare summary totals
    print(f"\n💰 Summary Totals Comparison:")
    excel_summary = excel_energy.get("summary", {})
    
    excel_total_cost = excel_summary.get("total_energy_cost", 0)
    excel_electrical_cost = excel_summary.get("total_electrical_cost", 0)
    
    print(f"   Excel Total Energy Cost: ${excel_total_cost:,.2f}")
    print(f"   Excel Electrical Cost: ${excel_electrical_cost:,.2f}")
    
    # Save detailed comparison
    comparison_results = {
        "excel_energy_data": excel_energy,
        "document_energy_data": doc_energy,
        "comparison_summary": {
            "common_energy_types": list(common_types),
            "excel_only_types": list(excel_only),
            "document_only_types": list(doc_only),
            "cost_matches": cost_matches,
            "cost_differences": cost_differences,
            "cost_match_rate_percent": cost_match_rate,
            "usage_matches": usage_matches,
            "usage_differences": usage_differences,
            "usage_match_rate_percent": usage_match_rate
        }
    }
    
    with open("energy_comparison_results.json", "w") as f:
        json.dump(comparison_results, f, indent=2, default=str)
    
    print(f"\n✅ Detailed energy comparison saved to: energy_comparison_results.json")
    
    return comparison_results

if __name__ == "__main__":
    compare_energy_data()
