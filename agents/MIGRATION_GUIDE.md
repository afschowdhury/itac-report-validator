# Migration Guide: v1.0 to v2.0

## Overview

The agents package has been restructured to follow Google ADK best practices with configuration-driven design. This guide helps you migrate from the old structure to the new one.

## What Changed

### Old Structure (v1.0)
```
agents/
├── __init__.py
├── summary_agent.py       # All code and config in one file
└── demo_summary_checker.py
```

### New Structure (v2.0)
```
agents/
├── __init__.py                    # Updated exports
├── summary_checker/               # NEW: Agent folder
│   ├── __init__.py               # Module exports
│   ├── agent.py                  # Agent implementation
│   └── config.toml               # All configuration
├── summary_agent.py               # Compatibility layer
└── demo_summary_checker.py        # Updated to use new structure
```

## Migration Steps

### Step 1: Update Imports (Recommended)

**Old Code:**
```python
from agents.summary_agent import (
    create_summary_checker_agent,
    check_all_ar_summaries,
    analyze_with_llm
)
```

**New Code:**
```python
from agents.summary_checker import (
    create_agent,
    check_all_ar_summaries,
    analyze_with_llm
)
```

### Step 2: Update Function Calls

**Old:**
```python
agent = create_summary_checker_agent(api_key=api_key)
```

**New:**
```python
agent = create_agent(api_key=api_key)
```

### Step 3: Configuration (Optional)

Instead of hardcoded values, you can now modify `agents/summary_checker/config.toml`:

```toml
[model]
name = "gemini-2.0-flash"
temperature = 0.7
max_tokens = 2048

[prompts]
system_instruction = """
Your custom prompt here...
"""
```

## Backward Compatibility

**Good News:** The old imports still work! They will show deprecation warnings but won't break your code.

```python
# This still works (with deprecation warning)
from agents.summary_agent import create_summary_checker_agent
agent = create_summary_checker_agent(api_key=api_key)
```

This gives you time to migrate gradually.

## New Features in v2.0

### 1. Configuration Files
All settings in `config.toml` - no code changes needed to adjust parameters.

### 2. Additional Tools
```python
from agents.summary_checker import (
    compare_ar_data,          # NEW: Compare AR with summary
    analyze_discrepancies,    # NEW: Pattern analysis
    get_agent_config          # NEW: Get configuration
)
```

### 3. ADK CLI Support
```bash
# Run agent interactively
adk run agents.summary_checker

# Launch dev UI
adk web

# Run as API server
adk api_server agents.summary_checker
```

### 4. Better Tool Integration
Tools are now proper `FunctionTool` instances with better documentation and type hints.

## Testing Your Migration

Run the test script to verify everything works:

```bash
python agents/test_new_structure.py
```

## Common Issues

### Import Error
**Problem:** `ModuleNotFoundError: No module named 'google.adk'`

**Solution:** Make sure `google-adk` is installed:
```bash
pip install google-adk tomli python-dotenv
```

### Config Not Found
**Problem:** `FileNotFoundError: config.toml not found`

**Solution:** Make sure you're running from the project root directory.

### Deprecation Warnings
**Problem:** Seeing `DeprecationWarning` messages

**Solution:** This is expected with old imports. Update to new imports to remove warnings:
```python
# Change this:
from agents.summary_agent import check_all_ar_summaries

# To this:
from agents.summary_checker import check_all_ar_summaries
```

## Need Help?

- Check `agents/README.md` for usage examples
- See `agents/ARCHITECTURE.md` for technical details
- Run `python agents/test_new_structure.py` to verify setup
- Look at `agents/demo_summary_checker.py` for complete example

## Rollback (If Needed)

If you need to temporarily rollback, the old `summary_agent.py` file contains all the original logic and will continue to work. However, we recommend migrating to v2.0 for better maintainability and future features.

