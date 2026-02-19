"""
Summary Checker Agent using Google ADK (Legacy Compatibility Layer)

DEPRECATED: This module is maintained for backward compatibility only.
Please use the new structure: from agents.summary_checker import create_agent

This agent validates AR (Assessment Recommendation) summaries by comparing
textual descriptions against extracted numerical data to identify inconsistencies.
"""

import warnings
from pathlib import Path
from typing import Any, Dict, List, Optional

from google import genai
from google.genai import types

from .summary_checker.agent import analyze_with_llm as _new_analyze_with_llm
from .summary_checker.agent import check_all_ar_summaries as _new_check_all_ar_summaries
from .summary_checker.agent import create_agent as _new_create_agent

# Import from new structure
from .summary_checker.agent import validate_ar_summary as _new_validate_ar_summary

# Deprecation warning
warnings.warn(
    "The 'summary_agent' module is deprecated. "
    "Please use 'from agents.summary_checker import create_agent' instead.",
    DeprecationWarning,
    stacklevel=2
)


def validate_ar_summary(
    ar_summary: str,
    ar_number: int,
    comparison_data: Dict[str, Any]
) -> Dict[str, Any]:
    """
    DEPRECATED: Use agents.summary_checker.validate_ar_summary instead.
    
    Validate an AR summary against its numerical data.
    
    Args:
        ar_summary: Textual summary of the AR
        ar_number: AR number being validated
        comparison_data: Data from compare_ar_with_summary function
        
    Returns:
        Dictionary containing validation results with identified issues
    """
    return _new_validate_ar_summary(ar_summary, ar_number, comparison_data)


def create_summary_checker_agent(api_key: Optional[str] = None, config_path: Optional[Path] = None):
    """
    DEPRECATED: Use agents.summary_checker.create_agent instead.
    
    Create and configure the AR Summary Checker Agent using Google ADK.
    
    Args:
        api_key: Google API key for authentication
        config_path: Path to config.toml file
        
    Returns:
        Configured LlmAgent for summary validation
    """
    return _new_create_agent(api_key=api_key, config_path=config_path)


def check_all_ar_summaries(
    ar_summaries: List[Dict[str, Any]],
    summary_recommendations: List[Dict[str, Any]],
    ar_data_list: List[Dict[str, Any]],
    api_key: Optional[str] = None,
    config_path: Optional[Path] = None
) -> List[Dict[str, Any]]:
    """
    DEPRECATED: Use agents.summary_checker.check_all_ar_summaries instead.
    
    Check all AR summaries against their numerical data.
    
    Args:
        ar_summaries: List of AR summaries from parse_ar_summaries
        summary_recommendations: List from get_recommended_summary_table_json
        ar_data_list: List of individual AR data from get_single_ar_summary_table
        api_key: Google API key
        config_path: Path to config file
        
    Returns:
        List of validation results for each AR
    """
    return _new_check_all_ar_summaries(
        ar_summaries,
        summary_recommendations,
        ar_data_list,
        api_key=api_key,
        config_path=config_path
    )


def analyze_with_llm(
    validation_results: List[Dict[str, Any]],
    api_key: Optional[str] = None,
    config_path: Optional[Path] = None
) -> str:
    """
    DEPRECATED: Use agents.summary_checker.analyze_with_llm instead.
    
    Use LLM to analyze validation results and provide insights.
    
    Args:
        validation_results: Results from check_all_ar_summaries
        api_key: Google API key
        config_path: Path to config file
        
    Returns:
        Analysis report as a string
    """
    return _new_analyze_with_llm(
        validation_results,
        api_key=api_key,
        config_path=config_path
    )


# Example usage
if __name__ == "__main__":
    import os

    from doc_extractor_utils import (
        get_recommended_summary_table_json,
        get_single_ar_summary_table,
        parse_ar_summaries,
    )
    
    # This is example code - actual usage would load real data
    print("AR Summary Checker Agent")
    print("=" * 50)
    print("\nThis agent validates AR summaries against numerical data.")
    print("\nUsage:")
    print("1. Load AR summaries using parse_ar_summaries()")
    print("2. Load recommendation data using get_recommended_summary_table_json()")
    print("3. Load individual AR data using get_single_ar_summary_table()")
    print("4. Run check_all_ar_summaries() to validate")
    print("5. Optionally use analyze_with_llm() for detailed insights")
    
    # Example structure
    example_usage = """
    # Example:
    from agents.summary_agent import check_all_ar_summaries, analyze_with_llm
    
    # Load your data
    ar_summaries = parse_ar_summaries(ar_summary_html)
    rec_summary = get_recommended_summary_table_json(rec_summary_html)
    ar_data_list = [get_single_ar_summary_table(ar_html) for ar_html in ar_htmls]
    
    # Check summaries
    results = check_all_ar_summaries(
        ar_summaries,
        rec_summary['recommendations'],
        ar_data_list,
        api_key=os.getenv('GOOGLE_API_KEY')
    )
    
    # Get LLM analysis
    analysis = analyze_with_llm(results, api_key=os.getenv('GOOGLE_API_KEY'))
    print(analysis)
    """
    
    print(example_usage)

