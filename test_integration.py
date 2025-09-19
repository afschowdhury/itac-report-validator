#!/usr/bin/env python3
"""
Simple test script showing how to use excel_keyinfo_extractor functions
without the argument parser.
"""

from excel_keyinfo_extractor import extract_all_structured_info
import json
from pathlib import Path

def test_extraction():
    """Test the enhanced extractor functions."""
    
    # Path to test file
    test_file = "docs/report1/IACAssessmentTemplate.v2.1.xlsx"
    
    if not Path(test_file).exists():
        print(f"Test file not found: {test_file}")
        return
    
    print("🔍 Testing excel_keyinfo_extractor integration...")
    
    try:
        # Extract all structured data
        data = extract_all_structured_info(test_file)
        
        # Show what we extracted
        print(f"\n📋 General Info Fields: {len(data['general_info'])}")
        print(f"⚡ Energy Sources: {len(data['energy_waste_info']['energy_sources'])}")
        print(f"💡 Recommendations: {len(data['recommendation_info']['recommendations'])}")
        
        # Show some key values
        general = data["general_info"]
        energy = data["energy_waste_info"]["summary"]
        
        print(f"\n🏭 Company Details:")
        print(f"   Product: {general.get('Principle Product', 'N/A')}")
        print(f"   Employees: {general.get('# of Employees', 'N/A')}")
        print(f"   Annual Sales: ${general.get('Annual Sales ($)', 0):,.0f}")
        
        print(f"\n💰 Energy Summary:")
        print(f"   Total Energy Cost: ${energy.get('total_energy_cost', 0):,.2f}")
        print(f"   Electrical Cost: ${energy.get('total_electrical_cost', 0):,.2f}")
        
        # Save result
        output_file = "integration_test_result.json"
        with open(output_file, "w") as f:
            json.dump(data, f, indent=2, default=str)
        
        print(f"\n✅ Success! Data saved to: {output_file}")
        return data
        
    except Exception as e:
        print(f"❌ Error: {e}")
        return None

if __name__ == "__main__":
    test_extraction()
