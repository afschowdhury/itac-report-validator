"""
AI Agents for ITAC Report Validation

This package provides ADK-based agents for validating ITAC reports.
"""

# New ADK-based structure
from .summary_checker import (
    create_agent,
    get_agent_config,
    check_all_ar_summaries,
    analyze_with_llm,
    validate_ar_summary,
    compare_ar_data,
    analyze_discrepancies
)

# Backward compatibility - old imports still work
from .summary_agent import (
    create_summary_checker_agent,
    check_all_ar_summaries as check_all_ar_summaries_legacy,
    analyze_with_llm as analyze_with_llm_legacy,
    validate_ar_summary as validate_ar_summary_legacy
)

__version__ = "2.0.0"

# Primary exports (new ADK structure)
__all__ = [
    # New ADK-based exports
    'create_agent',
    'get_agent_config',
    'check_all_ar_summaries',
    'analyze_with_llm',
    'validate_ar_summary',
    'compare_ar_data',
    'analyze_discrepancies',
    # Legacy exports for backward compatibility
    'create_summary_checker_agent',
    'check_all_ar_summaries_legacy',
    'analyze_with_llm_legacy',
    'validate_ar_summary_legacy',
]

