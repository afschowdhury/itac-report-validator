#!/usr/bin/env python
"""
Test script for the new ADK-based agent structure
"""

import sys
from pathlib import Path

# Add parent directory to path
sys.path.insert(0, str(Path(__file__).parent.parent))

def test_imports():
    """Test that all imports work correctly"""
    print("Testing imports...")
    
    try:
        from agents.summary_checker import (
            create_agent,
            get_agent_config,
            validate_ar_summary,
            compare_ar_data,
            analyze_discrepancies,
            check_all_ar_summaries,
            analyze_with_llm
        )
        print("✓ New imports work")
        
        # Test backward compatibility
        from agents.summary_agent import (
            create_summary_checker_agent,
            check_all_ar_summaries as legacy_check,
            validate_ar_summary as legacy_validate
        )
        print("✓ Legacy imports work (with deprecation warnings)")
        
    except ImportError as e:
        print(f"✗ Import failed: {e}")
        return False
    
    return True


def test_config():
    """Test configuration loading"""
    print("\nTesting configuration...")
    
    try:
        from agents.summary_checker import get_agent_config
        
        config = get_agent_config()
        
        # Check all required sections
        assert 'agent' in config, "Missing 'agent' section"
        assert 'model' in config, "Missing 'model' section"
        assert 'prompts' in config, "Missing 'prompts' section"
        assert 'tools' in config, "Missing 'tools' section"
        
        print(f"✓ Config loaded successfully")
        print(f"  - Agent: {config['agent']['name']}")
        print(f"  - Model: {config['model']['name']}")
        print(f"  - Temperature: {config['model']['temperature']}")
        print(f"  - Tools: {', '.join(config['tools']['enabled'])}")
        
    except Exception as e:
        print(f"✗ Config test failed: {e}")
        return False
    
    return True


def test_tools():
    """Test tool functions"""
    print("\nTesting tool functions...")
    
    try:
        from agents.summary_checker import (
            validate_ar_summary,
            compare_ar_data,
            analyze_discrepancies
        )
        
        # Test validate_ar_summary
        test_comparison_data = {
            'matches': [
                {'field': 'electricity_savings_kwh', 'ar_value': 100000},
                {'field': 'cost_savings', 'ar_value': 5000}
            ],
            'differences': [
                {
                    'field': 'implementation_cost',
                    'ar_value': 10000,
                    'summary_value': 9000,
                    'difference': 1000
                }
            ]
        }
        
        result = validate_ar_summary(
            ar_summary="Test summary for AR 1",
            ar_number=1,
            comparison_data=test_comparison_data
        )
        
        assert 'ar_number' in result
        assert 'has_differences' in result
        assert result['has_differences'] == True
        assert result['total_matches'] == 2
        assert result['total_differences'] == 1
        
        print("✓ validate_ar_summary works")
        
        # Test analyze_discrepancies
        test_validation_results = [
            {
                'validation': {'has_differences': True},
                'comparison': {
                    'differences': [
                        {'field': 'cost_savings'},
                        {'field': 'implementation_cost'}
                    ]
                }
            },
            {
                'validation': {'has_differences': False},
                'comparison': {'differences': []}
            }
        ]
        
        analysis = analyze_discrepancies(test_validation_results)
        
        assert 'total_ars_analyzed' in analysis
        assert analysis['total_ars_analyzed'] == 2
        assert analysis['ars_with_discrepancies'] == 1
        
        print("✓ analyze_discrepancies works")
        
    except Exception as e:
        print(f"✗ Tool test failed: {e}")
        import traceback
        traceback.print_exc()
        return False
    
    return True


def test_structure():
    """Test that folder structure is correct"""
    print("\nTesting folder structure...")
    
    base_path = Path(__file__).parent
    
    required_files = [
        base_path / "summary_checker" / "__init__.py",
        base_path / "summary_checker" / "agent.py",
        base_path / "summary_checker" / "config.toml",
        base_path / "__init__.py",
        base_path / "summary_agent.py",
        base_path / "demo_summary_checker.py",
        base_path / "README.md",
        base_path / "ARCHITECTURE.md"
    ]
    
    all_exist = True
    for file_path in required_files:
        if file_path.exists():
            print(f"✓ {file_path.relative_to(base_path.parent)}")
        else:
            print(f"✗ Missing: {file_path.relative_to(base_path.parent)}")
            all_exist = False
    
    return all_exist


def main():
    """Run all tests"""
    print("=" * 70)
    print("TESTING NEW ADK AGENT STRUCTURE")
    print("=" * 70)
    print()
    
    tests = [
        ("Imports", test_imports),
        ("Configuration", test_config),
        ("Tools", test_tools),
        ("Folder Structure", test_structure)
    ]
    
    results = []
    for name, test_func in tests:
        try:
            result = test_func()
            results.append((name, result))
        except Exception as e:
            print(f"\n✗ {name} test crashed: {e}")
            import traceback
            traceback.print_exc()
            results.append((name, False))
    
    print("\n" + "=" * 70)
    print("TEST SUMMARY")
    print("=" * 70)
    
    for name, result in results:
        status = "✓ PASSED" if result else "✗ FAILED"
        print(f"{name:.<50} {status}")
    
    all_passed = all(result for _, result in results)
    
    print("\n" + "=" * 70)
    if all_passed:
        print("✓ ALL TESTS PASSED")
        return 0
    else:
        print("✗ SOME TESTS FAILED")
        return 1


if __name__ == "__main__":
    import warnings
    warnings.filterwarnings('ignore', category=DeprecationWarning)
    sys.exit(main())

