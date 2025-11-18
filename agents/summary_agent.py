"""
Summary Checker Agent using Google ADK

This agent validates AR (Assessment Recommendation) summaries by comparing
textual descriptions against extracted numerical data to identify inconsistencies.
"""

from typing import Any, Dict, List
from google import genai
from google.genai import types


def validate_ar_summary(
    ar_summary: str,
    ar_number: int,
    comparison_data: Dict[str, Any]
) -> Dict[str, Any]:
    """
    Validate an AR summary against its numerical data.
    
    Args:
        ar_summary: Textual summary of the AR
        ar_number: AR number being validated
        comparison_data: Data from compare_ar_with_summary function
        
    Returns:
        Dictionary containing validation results with identified issues
    """
    # Extract key metrics from comparison data
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
        'total_differences': len(differences)
    }


def create_summary_checker_agent(api_key: str = None):
    """
    Create and configure the AR Summary Checker Agent using Google ADK.
    
    Args:
        api_key: Google API key for authentication
        
    Returns:
        Configured LlmAgent for summary validation
    """
    from google.adk.agents import LlmAgent
    from google.adk.tools import FunctionTool
    
    # Configure the client
    client = genai.Client(api_key=api_key)
    
    # Create the validation tool
    validation_tool = FunctionTool(validate_ar_summary)
    
    # Create the LLM agent
    agent = LlmAgent(
        name="ar_summary_validator",
        model="gemini-2.0-flash",
        description="Validates Assessment Recommendation summaries against numerical data",
        instruction="""
        You are an expert validator for Industrial Assessment Center (IAC) reports.
        Your task is to review AR (Assessment Recommendation) summaries and identify
        any inconsistencies between the textual description and the numerical data.
        
        For each AR summary, you should:
        1. Carefully read the textual summary
        2. Compare mentioned savings, costs, and metrics with the extracted numerical data
        3. Identify any discrepancies, contradictions, or missing information
        4. Flag potential data entry errors or calculation mistakes
        5. Provide clear, actionable feedback on what needs to be corrected
        
        Focus on:
        - Energy savings (kWh/yr, MMBtu/yr)
        - Cost savings ($/yr)
        - Implementation costs ($)
        - Payback periods (years)
        - CO2 reductions (tons/yr)
        - Demand savings (kW/yr)
        
        Be thorough but concise. Prioritize identifying actual errors over minor formatting issues.
        """,
        tools=[validation_tool]
    )
    
    return agent


def check_all_ar_summaries(
    ar_summaries: List[Dict[str, Any]],
    summary_recommendations: List[Dict[str, Any]],
    ar_data_list: List[Dict[str, Any]],
    api_key: str = None
) -> List[Dict[str, Any]]:
    """
    Check all AR summaries against their numerical data.
    
    Args:
        ar_summaries: List of AR summaries from parse_ar_summaries
        summary_recommendations: List from get_recommended_summary_table_json
        ar_data_list: List of individual AR data from get_single_ar_summary_table
        api_key: Google API key
        
    Returns:
        List of validation results for each AR
    """
    from doc_extractor_utils import compare_ar_with_summary
    
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
        comparison = compare_ar_with_summary(ar_data, summary_recommendations)
        
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
    api_key: str = None
) -> str:
    """
    Use LLM to analyze validation results and provide insights.
    
    Args:
        validation_results: Results from check_all_ar_summaries
        api_key: Google API key
        
    Returns:
        Analysis report as a string
    """
    client = genai.Client(api_key=api_key)
    
    # Build the analysis prompt
    prompt = """
    Please analyze the following AR summary validation results and provide a comprehensive report.
    
    Identify:
    1. ARs with data inconsistencies
    2. Common patterns in discrepancies
    3. Severity of issues (critical vs minor)
    4. Recommendations for corrections
    
    Validation Results:
    """
    
    for result in validation_results:
        if result.get('status') == 'error':
            prompt += f"\n\nAR {result['ar_number']}: ERROR - {result['message']}"
            continue
        
        validation = result.get('validation', {})
        prompt += f"\n\nAR {validation.get('ar_number')}:"
        prompt += f"\n{validation.get('context', '')}"
    
    # Generate analysis using Gemini
    response = client.models.generate_content(
        model="gemini-2.5-pro",
        contents=prompt
    )
    
    return response.text


# Example usage
if __name__ == "__main__":
    import os
    from doc_extractor_utils import (
        parse_ar_summaries,
        get_recommended_summary_table_json,
        get_single_ar_summary_table
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

