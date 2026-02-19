# AI Agents for ITAC Report Validation

This directory contains AI agents built with Google's Agent Development Kit (ADK) for validating ITAC (Industrial Assessment Center) reports.

## Overview

The **Summary Checker Agent** validates Assessment Recommendation (AR) summaries by comparing textual descriptions against extracted numerical data to identify inconsistencies, errors, and discrepancies.

### New Structure (v2.0)

The agents are now organized using ADK best practices with separate configuration files:

```
agents/
├── __init__.py                    # Package exports
├── summary_checker/               # AR Summary Checker Agent (NEW)
│   ├── __init__.py               # Module exports
│   ├── agent.py                  # ADK agent implementation
│   └── config.toml               # Configuration (prompts, model params)
├── summary_agent.py               # Legacy compatibility layer
└── demo_summary_checker.py        # Demo script
```

## Features

- ✅ **Automated Validation**: Compares AR summary text with numerical data
- 🔍 **Discrepancy Detection**: Identifies mismatches between text and numbers
- 🤖 **AI-Powered Analysis**: Uses Gemini LLM to provide intelligent insights
- 📊 **Comprehensive Reports**: Generates detailed validation reports
- 🔧 **Modular Design**: Built on Google ADK for extensibility

## Architecture

The agent uses Google ADK's **LLM Agent** architecture with TOML-based configuration:

```
┌─────────────────────────────────────────────────┐
│   AR Summary Checker Agent                      │
│   (LlmAgent with Gemini 2.0 Flash)             │
│   Config: agents/summary_checker/config.toml    │
└─────────────────┬───────────────────────────────┘
                  │
                  ├─► validate_ar_summary() - FunctionTool
                  │   Validates summary against data
                  │
                  ├─► compare_ar_data() - FunctionTool
                  │   Compares AR with summary table
                  │
                  └─► analyze_discrepancies() - FunctionTool
                      Finds patterns in errors
```

### Configuration File (config.toml)

All agent parameters are stored in `agents/summary_checker/config.toml`:

- **Agent settings**: name, description, version
- **Model parameters**: temperature, max_tokens, top_p, top_k
- **Prompts**: system instructions, validation prompts
- **Tools**: enabled tools and their configurations
- **Validation settings**: thresholds, required metrics
- **Output preferences**: format, verbosity

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

### Basic Usage (New ADK Structure)

```python
from agents.summary_checker import (
    create_agent,
    check_all_ar_summaries,
    analyze_with_llm,
    get_agent_config
)
from doc_extractor_utils import (
    parse_ar_summaries,
    get_recommended_summary_table_json,
    get_single_ar_summary_table
)

# Get agent configuration
config = get_agent_config()
print(f"Using model: {config['model']['name']}")

# Create the agent (optional - for custom usage)
agent = create_agent(api_key=os.getenv('GOOGLE_API_KEY'))

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

### Legacy Usage (Backward Compatible)

```python
# Old imports still work with deprecation warnings
from agents.summary_agent import check_all_ar_summaries, analyze_with_llm
# ... rest of code remains the same
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

### Creating and Customizing Agents

```python
from agents.summary_checker import create_agent
from pathlib import Path

# Create agent with default configuration
agent = create_agent(api_key='your-api-key')

# Create agent with custom config file
custom_config_path = Path('path/to/custom_config.toml')
agent = create_agent(
    api_key='your-api-key',
    config_path=custom_config_path
)

# Override specific parameters
agent = create_agent(
    api_key='your-api-key',
    model='gemini-2.0-flash-exp',  # Use experimental model
    name='custom_validator'
)
```

### Editing Configuration

Modify `agents/summary_checker/config.toml` to customize:

```toml
[model]
name = "gemini-2.0-flash"
temperature = 0.7    # Adjust creativity (0.0-1.0)
max_tokens = 2048    # Max response length

[prompts]
system_instruction = """
Your custom prompt here...
"""
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
├── __init__.py                    # Package exports (v2.0)
├── summary_checker/               # AR Summary Checker Agent
│   ├── __init__.py               # Module exports
│   ├── agent.py                  # ADK agent implementation
│   └── config.toml               # All configurations
├── summary_agent.py               # Legacy compatibility layer
├── demo_summary_checker.py        # Demo script (updated for v2.0)
├── README.md                      # This file
├── ARCHITECTURE.md                # Detailed architecture docs
└── SUMMARY.md                     # Summary documentation
```

## Dependencies

- `google-adk`: Google Agent Development Kit (v0.1.0+)
- `google-genai`: Google Generative AI SDK
- `tomli`: TOML configuration parsing (Python < 3.11)
- `beautifulsoup4`: HTML parsing
- Other dependencies in `requirements.txt`

Install all dependencies:
```bash
pip install -r requirements.txt
```

Or install just the agent dependencies:
```bash
pip install google-adk google-genai tomli
```

## References

- [Google ADK Documentation](https://google.github.io/adk-docs/)
- [LLM Agents Guide](https://google.github.io/adk-docs/agents/llm-agents/)
- [Workflow Agents](https://google.github.io/adk-docs/agents/workflow-agents/)
- [Function Tools](https://google.github.io/adk-docs/tools/function-tools/)

## Running with ADK CLI

The agent can be run using ADK's command-line interface:

```bash
# Run agent in terminal (interactive)
cd /path/to/itac-report-validator
adk run agents.summary_checker

# Run agent with dev UI (browser-based)
adk web

# Run as API server
adk api_server agents.summary_checker
```

## Migration Guide

### From v1.0 to v2.0

If you're using the old structure, here's how to migrate:

**Old (v1.0):**
```python
from agents.summary_agent import create_summary_checker_agent
agent = create_summary_checker_agent(api_key=key)
```

**New (v2.0):**
```python
from agents.summary_checker import create_agent
agent = create_agent(api_key=key)
```

The old imports will continue to work but show deprecation warnings.

## Future Enhancements

- [ ] Add energy validator agent (separate folder)
- [ ] Add financial validator agent (separate folder)
- [ ] Implement workflow agents (Sequential, Loop) for orchestration
- [ ] Add custom tools for domain-specific validation rules
- [ ] Integration with report generation system
- [ ] Real-time validation during document extraction
- [ ] Support for batch processing multiple reports
- [ ] Add agent templates for easy agent creation

## License

Part of the ITAC Report Validator project.

