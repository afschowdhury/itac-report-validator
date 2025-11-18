# AR Summary Checker Agent

This directory contains AI agents built with Google's Agent Development Kit (ADK) for validating ITAC (Industrial Assessment Center) reports.

## Overview

The **Summary Checker Agent** validates Assessment Recommendation (AR) summaries by comparing textual descriptions against extracted numerical data to identify inconsistencies, errors, and discrepancies.

## Features

- ✅ **Automated Validation**: Compares AR summary text with numerical data
- 🔍 **Discrepancy Detection**: Identifies mismatches between text and numbers
- 🤖 **AI-Powered Analysis**: Uses Gemini LLM to provide intelligent insights
- 📊 **Comprehensive Reports**: Generates detailed validation reports
- 🔧 **Modular Design**: Built on Google ADK for extensibility

## Architecture

The agent uses Google ADK's **LLM Agent** architecture:

```
┌─────────────────────────────────────┐
│   AR Summary Checker Agent          │
│   (LlmAgent with Gemini)            │
└─────────────────┬───────────────────┘
                  │
                  ├─► validate_ar_summary() - Function Tool
                  │
                  ├─► parse_ar_summaries() - Data Parser
                  │
                  └─► compare_ar_with_summary() - Data Comparator
```

### Agent Types (from Google ADK)

1. **LLM Agents**: Use Large Language Models for dynamic reasoning (used here)
2. **Workflow Agents**: Control execution flow
   - Sequential Agents
   - Loop Agents
   - Parallel Agents
3. **Custom Agents**: Fully customizable agents

## Installation

1. Install dependencies:
```bash
pip install -r requirements.txt
```

2. Set up your Google API key:
```bash
export GOOGLE_API_KEY='your-api-key-here'
```

## Usage

### Basic Usage

```python
from agents.summary_agent import check_all_ar_summaries, analyze_with_llm
from doc_extractor_utils import (
    parse_ar_summaries,
    get_recommended_summary_table_json,
    get_single_ar_summary_table
)

# Load your extracted HTML data
ar_summaries = parse_ar_summaries(ar_summary_html)
rec_summary = get_recommended_summary_table_json(rec_summary_html)
ar_data_list = [get_single_ar_summary_table(ar_html) for ar_html in ar_htmls]

# Check all summaries
results = check_all_ar_summaries(
    ar_summaries,
    rec_summary['recommendations'],
    ar_data_list,
    api_key=os.getenv('GOOGLE_API_KEY')
)

# Get AI-powered analysis
analysis = analyze_with_llm(results, api_key=os.getenv('GOOGLE_API_KEY'))
print(analysis)
```

### Running the Demo

```bash
# Make sure you have extracted HTML files in EXTRACTED_HTML/
python agents/demo_summary_checker.py
```

The demo will:
1. Load AR summaries from extracted HTML files
2. Validate each AR summary against numerical data
3. Generate a comprehensive validation report
4. Use Gemini LLM to provide intelligent analysis
5. Save results to JSON files

## How It Works

### 1. Data Loading
- Loads AR summaries using `parse_ar_summaries()`
- Loads recommendation data using `get_recommended_summary_table_json()`
- Loads individual AR data using `get_single_ar_summary_table()`

### 2. Validation Process
For each AR:
- Extracts numerical data from AR savings summary table
- Compares with corresponding row in recommendation summary table
- Identifies matches and differences
- Validates textual summary against numerical data

### 3. AI Analysis
- Uses Gemini LLM to analyze all validation results
- Identifies patterns and common issues
- Categorizes severity of discrepancies
- Provides actionable recommendations

## Validation Metrics

The agent checks consistency for:

- **Energy Savings**: kWh/yr, MMBtu/yr
- **Cost Savings**: $/yr (energy, demand, admin, propane)
- **Implementation Cost**: $
- **Payback Period**: years
- **CO2 Reduction**: tons/yr
- **Demand Savings**: kW/yr

## Output

### Validation Results
```json
{
  "ar_number": 1,
  "summary": "AR text...",
  "validation": {
    "has_differences": false,
    "total_matches": 8,
    "total_differences": 0
  },
  "comparison": {
    "matches": [...],
    "differences": [...]
  }
}
```

### AI Analysis Report
The LLM provides a comprehensive analysis including:
- ARs with critical discrepancies
- Common patterns in errors
- Severity assessment
- Recommendations for corrections

## Advanced Usage

### Creating Custom Agents

```python
from agents.summary_agent import create_summary_checker_agent

# Create the agent
agent = create_summary_checker_agent(api_key='your-api-key')

# The agent is configured with:
# - Name: "ar_summary_validator"
# - Model: "gemini-2.0-flash"
# - Custom instructions for IAC report validation
# - Validation tools
```

### Custom Validation Logic

```python
from agents.summary_agent import validate_ar_summary

# Direct validation of a single AR
result = validate_ar_summary(
    ar_summary="AR No. 1 - Replace inefficient motors...",
    ar_number=1,
    comparison_data=comparison_results
)
```

## Project Structure

```
agents/
├── __init__.py                 # Package initialization
├── summary_agent.py            # Main agent implementation
├── demo_summary_checker.py     # Demo script
└── README.md                   # This file
```

## Dependencies

- `google-adk`: Google Agent Development Kit
- `google-genai`: Google Generative AI SDK
- `beautifulsoup4`: HTML parsing
- Other dependencies in `requirements.txt`

## References

- [Google ADK Documentation](https://google.github.io/adk-docs/)
- [LLM Agents Guide](https://google.github.io/adk-docs/agents/llm-agents/)
- [Workflow Agents](https://google.github.io/adk-docs/agents/workflow-agents/)
- [Function Tools](https://google.github.io/adk-docs/tools/function-tools/)

## Future Enhancements

- [ ] Add multi-agent workflow for hierarchical validation
- [ ] Implement sequential agent for step-by-step validation pipeline
- [ ] Add custom tools for domain-specific validation rules
- [ ] Integration with report generation system
- [ ] Real-time validation during document extraction
- [ ] Support for batch processing multiple reports

## License

Part of the ITAC Report Validator project.

