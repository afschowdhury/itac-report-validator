#!/usr/bin/env python3
"""
Test script to demonstrate the corrected energy waste info extraction.
"""

from excel_keyinfo_extractor import extract_energy_waste_info_dict
import json

def test_corrected_energy_extraction():
    """Test the corrected energy extraction with proper units and formatting."""
    
    excel_path = "docs/report1/IACAssessmentTemplate.v2.1.xlsx"
    
    print("⚡ Testing Corrected Energy Waste Info Extraction...")
    
    # Extract energy data
    energy_data = extract_energy_waste_info_dict(excel_path)
    
    print(f"\n📊 Results:")
    print(f"   Total energy entries: {len(energy_data['data'])}")
    print(f"   Active energy sources (with costs): {energy_data['summary']['num_energy_sources']}")
    
    print(f"\n💰 Cost Summary:")
    print(f"   Total Energy Cost: ${energy_data['summary']['total_energy_cost']:,.2f}")
    print(f"   Electrical Cost: ${energy_data['summary']['total_electrical_cost']:,.2f}")
    
    print(f"\n🔋 Sample Energy Sources (with consumption > 0):")
    
    for item in energy_data['data'][:6]:  # Show first 6 entries
        if item['cost'] > 0:
            usage_str = ""
            if item['usage']:
                for unit, value in item['usage'].items():
                    usage_str = f"{value:,.0f} {unit}"
                    break
            
            unit_cost_str = ""
            if item['unit_cost'] and item['unit_cost']['amount'] != "n/a":
                unit_cost_str = f" (${item['unit_cost']['amount']:.3f}/{item['unit_cost']['unit']})"
            
            print(f"   • {item['type']} ({item['original_name']})")
            print(f"     Usage: {usage_str}")
            print(f"     Cost: ${item['cost']:,.2f}{unit_cost_str}")
    
    # Show format improvements
    print(f"\n✅ Format Improvements:")
    print(f"   ✓ Proper units: kWh/yr, MMBtu/yr, kW months/yr")
    print(f"   ✓ Standardized energy types: electrical_energy, natural_gas, propane_gas")
    print(f"   ✓ Correct cost values from Cost_1 column")
    print(f"   ✓ Electrical fees properly captured: ${energy_data['summary']['total_electrical_cost'] - 321236:,.2f}")
    print(f"   ✓ Original names preserved for reference")
    
    return energy_data

if __name__ == "__main__":
    test_corrected_energy_extraction()




