"""
AR Summary Checker Agent - ADK Implementation

This module implements the AR Summary Validation agent using Google's Agent Development Kit.
"""

import os
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional
import tomli

from google.adk.agents import LlmAgent
from google.adk.tools import FunctionTool
from google import genai
from dotenv import load_dotenv

from icecream import ic
ic.configureOutput(includeContext=True, prefix='DEBUG: ')


load_dotenv()

# Add project root to path to import doc_extractor_utils
PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))


def load_config(config_path: Optional[Path] = None) -> Dict[str, Any]:
    """
    Load agent configuration from TOML file.
    
    Args:
        config_path: Path to config.toml file. If None, uses default location.
        
    Returns:
        Dictionary containing configuration
    """
    if config_path is None:
        config_path = Path(__file__).parent / "config.toml"
    
    with open(config_path, 'rb') as f:
        config = tomli.load(f)
    
    return config


def get_agent_config(config_path: Optional[Path] = None) -> Dict[str, Any]:
    """
    Get agent configuration.
    
    Args:
        config_path: Optional path to config file
        
    Returns:
        Configuration dictionary
    """
    return load_config(config_path)


# Tool Functions
def validate_ar_summary(
    ar_summary: str,
    ar_number: int,
    comparison_data: Dict[str, Any]
) -> Dict[str, Any]:
    """
    Validate an AR summary against its numerical data.
    
    This is a tool that can be called by the LLM agent to validate
    Assessment Recommendation summaries.
    
    Args:
        ar_summary: Textual summary of the AR
        ar_number: AR number being validated
        comparison_data: Data from compare_ar_with_summary function
        
    Returns:
        Dictionary containing validation results with identified issues
    """
    # Extract key metrics from comparison data
    ic(comparison_data)
    matches = comparison_data.get('matches', [])
    differences = comparison_data.get('differences', [])
    
    # Build context for validation
    context = f"""
    AR Number: {ar_number}
    
    Summary Text:
    {ar_summary}
    
    Numerical Data Matches ({len(matches)}):
    """
    
    for match in matches:
        field = match.get('field', 'unknown')
        value = match.get('ar_value', 'N/A')
        context += f"\n  - {field}: {value}"
    
    if differences:
        context += f"\n\nData Inconsistencies Found ({len(differences)}):"
        for diff in differences:
            field = diff.get('field', 'unknown')
            ar_val = diff.get('ar_value', 'N/A')
            summary_val = diff.get('summary_value', 'N/A')
            difference = diff.get('difference', 'N/A')
            context += f"\n  - {field}: AR={ar_val}, Summary={summary_val}, Diff={difference}"
    
    return {
        'ar_number': ar_number,
        'context': context,
        'has_differences': len(differences) > 0,
        'total_matches': len(matches),
        'total_differences': len(differences),
        'summary_text': ar_summary,
        'matches': matches,
        'differences': differences
    }


def compare_ar_data(
    ar_data: Dict[str, Any],
    summary_recommendations: List[Dict[str, Any]]
) -> Dict[str, Any]:
    """
    Compare AR data with summary table recommendations.
    
    Tool for comparing individual AR data against the summary table.
    
    Args:
        ar_data: Individual AR data from get_single_ar_summary_table
        summary_recommendations: List of recommendations from summary table
        
    Returns:
        Comparison results with matches and differences
    """
    try:
        from doc_extractor_utils import compare_ar_with_summary
        return compare_ar_with_summary(ar_data, summary_recommendations)
    except Exception as e:
        return {
            'error': f"Failed to compare AR data: {str(e)}",
            'matches': [],
            'differences': []
        }


def analyze_discrepancies(
    validation_results: List[Dict[str, Any]]
) -> Dict[str, Any]:
    """
    Analyze patterns in data discrepancies across multiple ARs.
    
    Tool for identifying common patterns and issues across multiple AR validations.
    
    Args:
        validation_results: List of validation results from multiple ARs
        
    Returns:
        Analysis of common patterns and issues
    """
    total_ars = len(validation_results)
    ars_with_issues = sum(1 for r in validation_results if r.get('validation', {}).get('has_differences', False))
    
    # Collect all difference types
    difference_fields = {}
    for result in validation_results:
        if 'comparison' in result:
            for diff in result['comparison'].get('differences', []):
                field = diff.get('field', 'unknown')
                difference_fields[field] = difference_fields.get(field, 0) + 1
    
    return {
        'total_ars_analyzed': total_ars,
        'ars_with_discrepancies': ars_with_issues,
        'discrepancy_rate': ars_with_issues / total_ars if total_ars > 0 else 0,
        'common_discrepancy_fields': difference_fields,
        'summary': f"Analyzed {total_ars} ARs, found discrepancies in {ars_with_issues} ({ars_with_issues/total_ars*100:.1f}%)" if total_ars > 0 else "No ARs analyzed"
    }


def create_agent(
    api_key: Optional[str] = None,
    config_path: Optional[Path] = None,
    **kwargs
) -> LlmAgent:
    """
    Create and configure the AR Summary Checker Agent.
    
    Args:
        api_key: Google API key for authentication. If None, uses environment variable.
        config_path: Path to config.toml file. If None, uses default location.
        **kwargs: Additional configuration overrides
        
    Returns:
        Configured LlmAgent instance
    """
    # Load configuration
    config = load_config(config_path)
    
    # Get API key from environment if not provided
    if api_key is None:
        api_key = os.getenv('GOOGLE_API_KEY')
    
    # Extract configuration sections
    agent_config = config.get('agent', {})
    model_config = config.get('model', {})
    prompts = config.get('prompts', {})
    tools_config = config.get('tools', {})
    
    # Override with kwargs
    agent_name = kwargs.get('name', agent_config.get('name', 'ar_summary_validator'))
    model_name = kwargs.get('model', model_config.get('name', 'gemini-2.0-flash'))
    description = kwargs.get('description', agent_config.get('description', ''))
    instruction = kwargs.get('instruction', prompts.get('system_instruction', ''))
    
    # Create function tools
    validation_tool = FunctionTool(validate_ar_summary)
    compare_tool = FunctionTool(compare_ar_data)
    analysis_tool = FunctionTool(analyze_discrepancies)
    
    # Create the LLM agent
    agent = LlmAgent(
        name=agent_name,
        model=model_name,
        description=description,
        instruction=instruction,
        tools=[validation_tool, compare_tool, analysis_tool]
    )
    
    return agent


def check_all_ar_summaries(
    ar_summaries: List[Dict[str, Any]],
    summary_recommendations: List[Dict[str, Any]],
    ar_data_list: List[Dict[str, Any]],
    api_key: Optional[str] = None,
    config_path: Optional[Path] = None
) -> List[Dict[str, Any]]:
    """
    Check all AR summaries against their numerical data.
    
    This function provides the same interface as the original implementation
    but uses the new ADK-based structure.
    
    Args:
        ar_summaries: List of AR summaries from parse_ar_summaries
        summary_recommendations: List from get_recommended_summary_table_json
        ar_data_list: List of individual AR data from get_single_ar_summary_table
        api_key: Google API key
        config_path: Path to config file
        
    Returns:
        List of validation results for each AR
    """
    results = []
    
    for ar_summary_obj in ar_summaries:
        ar_no = ar_summary_obj['ar_no']
        ar_summary_text = ar_summary_obj['ar_summary']
        
        # Find corresponding AR data
        ar_data = next(
            (ar for ar in ar_data_list if ar.get('ar_number') == ar_no),
            None
        )
        
        if not ar_data:
            results.append({
                'ar_number': ar_no,
                'status': 'error',
                'message': f'No AR data found for AR {ar_no}'
            })
            continue
        
        # Compare AR data with summary table
        comparison = compare_ar_data(ar_data, summary_recommendations)
        
        if 'error' in comparison:
            results.append({
                'ar_number': ar_no,
                'status': 'error',
                'message': comparison['error']
            })
            continue
        
        # Validate the summary
        validation = validate_ar_summary(
            ar_summary_text,
            ar_no,
            comparison
        )
        
        results.append({
            'ar_number': ar_no,
            'summary': ar_summary_text,
            'validation': validation,
            'comparison': comparison
        })
    
    return results


def analyze_with_llm(
    validation_results: List[Dict[str, Any]],
    api_key: Optional[str] = None,
    config_path: Optional[Path] = None
) -> str:
    """
    Use LLM to analyze validation results and provide insights.
    
    Args:
        validation_results: Results from check_all_ar_summaries
        api_key: Google API key
        config_path: Path to config file
        
    Returns:
        Analysis report as a string
    """
    # Load configuration
    config = load_config(config_path)
    model_config = config.get('model', {})
    prompts = config.get('prompts', {})

    
    
    # Get API key
    if api_key is None:
        api_key = os.getenv('GOOGLE_API_KEY')
    
    # Create client
    client = genai.Client(api_key=api_key)
    
    # Build the analysis prompt
    validation_prompt_template = prompts.get('validation_prompt', '')
    
    # Format validation results
    results_text = ""
    for result in validation_results:
        if result.get('status') == 'error':
            results_text += f"\n\nAR {result['ar_number']}: ERROR - {result['message']}"
            continue
        
        validation = result.get('validation', {})
        results_text += f"\n\nAR {validation.get('ar_number')}:"
        results_text += f"\n{validation.get('context', '')}"
    
    prompt = validation_prompt_template.format(validation_results=results_text)

    ic(prompt)
    
    # Get model name (use fallback for analysis if configured)
    model_name = model_config.get('fallback', {}).get('name', model_config.get('name', 'gemini-2.0-flash'))
    
    # Generate analysis using Gemini
    response = client.models.generate_content(
        model=model_name,
        contents=prompt
    )
    
    return response.text


# Backward compatibility exports
__all__ = [
    'create_agent',
    'get_agent_config',
    'validate_ar_summary',
    'compare_ar_data',
    'analyze_discrepancies',
    'check_all_ar_summaries',
    'analyze_with_llm'
]

