"""
Demo script showing how to use the AR Summary Checker Agent

This script demonstrates how to:
1. Load AR summaries and data from HTML extracts
2. Validate summaries against numerical data
3. Get AI-powered analysis of discrepancies
"""

import os
import json
from pathlib import Path

# Add parent directory to path to import modules
import sys
sys.path.insert(0, str(Path(__file__).parent.parent))

from doc_extractor_utils import (
    parse_ar_summaries,
    get_recommended_summary_table_json,
    get_single_ar_summary_table
)
from agents.summary_agent import (
    check_all_ar_summaries,
    analyze_with_llm
)


def load_html_file(filepath: str) -> str:
    """Load HTML content from file."""
    with open(filepath, 'r', encoding='utf-8') as f:
        return f.read()


def main():
    """Main demo function."""
    print("=" * 70)
    print("AR SUMMARY CHECKER AGENT - DEMO")
    print("=" * 70)
    print()
    
    # Configuration
    html_dir = Path("EXTRACTED_HTML")
    api_key = os.getenv('GOOGLE_API_KEY')
    
    if not api_key:
        print("⚠️  WARNING: GOOGLE_API_KEY environment variable not set.")
        print("   Set it to use LLM-powered analysis:")
        print("   export GOOGLE_API_KEY='your-api-key-here'")
        print()
    
    # Check if HTML files exist
    if not html_dir.exists():
        print(f"❌ Error: {html_dir} directory not found")
        print("   Please run document extraction first")
        return
    
    print("📂 Loading extracted HTML data...")
    print()
    
    # Load AR summary HTML
    ar_summary_file = html_dir / "ar_summary.html"
    if not ar_summary_file.exists():
        print(f"❌ Error: {ar_summary_file} not found")
        return
    
    ar_summary_html = load_html_file(ar_summary_file)
    print(f"✓ Loaded AR summaries from: {ar_summary_file}")
    
    # Load recommendation summary table HTML
    rec_summary_file = html_dir / "recommendation_summary_table.html"
    if not rec_summary_file.exists():
        print(f"❌ Error: {rec_summary_file} not found")
        return
    
    rec_summary_html = load_html_file(rec_summary_file)
    print(f"✓ Loaded recommendation summary from: {rec_summary_file}")
    
    # Parse AR summaries
    print("\n📝 Parsing AR summaries...")
    ar_summaries = parse_ar_summaries(ar_summary_html)
    print(f"✓ Found {len(ar_summaries)} AR summaries")
    
    # Parse recommendation summary table
    print("\n📊 Parsing recommendation summary table...")
    rec_summary_data = get_recommended_summary_table_json(rec_summary_html)
    recommendations = rec_summary_data.get('recommendations', [])
    print(f"✓ Found {len(recommendations)} recommendations in summary table")
    
    # Load individual AR HTMLs
    print("\n📋 Loading individual AR data...")
    ar_data_list = []
    ar_files = sorted(html_dir.glob("AR_*.html"))
    
    for ar_file in ar_files:
        ar_html = load_html_file(ar_file)
        ar_data = get_single_ar_summary_table(ar_html)
        if ar_data.get('ar_number'):
            ar_data_list.append(ar_data)
            print(f"  ✓ Loaded AR {ar_data['ar_number']} from {ar_file.name}")
    
    print(f"\n✓ Total ARs loaded: {len(ar_data_list)}")
    
    # Validate summaries
    print("\n" + "=" * 70)
    print("VALIDATING AR SUMMARIES")
    print("=" * 70)
    print()
    
    validation_results = check_all_ar_summaries(
        ar_summaries,
        recommendations,
        ar_data_list,
        api_key=api_key
    )
    
    # Display results
    print(f"✓ Validated {len(validation_results)} AR summaries")
    print()
    
    for result in validation_results:
        ar_num = result.get('ar_number')
        
        if result.get('status') == 'error':
            print(f"❌ AR {ar_num}: {result.get('message')}")
            continue
        
        validation = result.get('validation', {})
        comparison = result.get('comparison', {})
        
        has_diffs = validation.get('has_differences', False)
        total_matches = validation.get('total_matches', 0)
        total_diffs = validation.get('total_differences', 0)
        
        status_icon = "⚠️ " if has_diffs else "✓"
        print(f"{status_icon} AR {ar_num}:")
        print(f"   Matches: {total_matches}, Differences: {total_diffs}")
        
        if has_diffs:
            print(f"   Discrepancies found:")
            for diff in comparison.get('differences', []):
                field = diff.get('field')
                ar_val = diff.get('ar_value')
                summary_val = diff.get('summary_value')
                print(f"     - {field}: AR={ar_val}, Summary={summary_val}")
        print()
    
    # Save results to JSON
    output_file = Path("ar_summary_validation_results.json")
    with open(output_file, 'w') as f:
        json.dump(validation_results, f, indent=2)
    print(f"💾 Results saved to: {output_file}")
    print()
    
    # LLM Analysis (if API key is available)
    if api_key and validation_results:
        print("=" * 70)
        print("AI-POWERED ANALYSIS")
        print("=" * 70)
        print()
        print("🤖 Generating comprehensive analysis using Gemini...")
        print()
        
        try:
            analysis = analyze_with_llm(validation_results, api_key=api_key)
            print(analysis)
            print()
            
            # Save analysis
            analysis_file = Path("ar_summary_analysis.txt")
            with open(analysis_file, 'w') as f:
                f.write(analysis)
            print(f"💾 Analysis saved to: {analysis_file}")
            
        except Exception as e:
            print(f"❌ Error during LLM analysis: {e}")
    else:
        print("ℹ️  Skipping LLM analysis (no API key provided)")
    
    print()
    print("=" * 70)
    print("DEMO COMPLETE")
    print("=" * 70)


if __name__ == "__main__":
    main()

