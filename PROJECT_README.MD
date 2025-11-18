# ITAC Report Validator

  The ITAC Report Validator is a Flask web app that cross-checks Industrial Assessment Center (IAC) assessment
  reports against their companion Excel templates. It parses both inputs, normalizes the data, compares
  key fields with intelligent tolerance handling, and presents mismatches in a browser-friendly layout. The
  codebase also exposes reusable extractors and helper utilities for standalone processing or automation.

  ---

  ## Key Capabilities

  - **Dual-pipeline ingestion** for DOCX reports and IAC Excel templates.
  - **Structured data extraction** of general facility info, energy usage, and assessment recommendations.
  - **Smart validation engine** with configurable numeric tolerances and rich mismatch metadata.
  - **API-first design** enabling browser, CLI, or automated batch comparisons.
  - **Agent helpers** to sanity-check AR summaries with LLM tooling (optional Google Gemini integration).

  ---

  ## Quick Start

  > Requires Python 3.10+ with `pip`.

  1. **Install dependencies**
     ```bash
     pip install -r requirements.txt
     ```

  2. **Run the web app**
     ```bash
     python app.py
     ```
     Open `http://localhost:5000` and upload one DOCX report and its matching Excel template.

  3. **Use the JSON API**
     ```bash
     curl -X POST http://localhost:5000/api/compare \
       -F "docx_file=@/path/to/report.docx" \
       -F "excel_file=@/path/to/template.xlsx"
     ```

  ---

  ## How It Works

  1. **DOCX extraction (`document_extractor.py`)**
     - Converts key sections to HTML or JSON.
     - Pulls structured values (e.g., SIC, NAICS, production metrics, AR tables).
     - Normalizes energy usage with unit-aware parsing.

  2. **Excel extraction (`excel_keyinfo_extractor.py`)**
     - Scans sheets for key–value pairs and tabular data.
    - Standardizes general info, energy/waste metrics, and recommendation data.

  3. **Comparison (`app.py`)**
     - Harmonizes field names between sources.
     - Compares strings (case-insensitive) and numerics (1% relative tolerance by default).
     - Aggregates validation summaries for the UI/API.

  4. **Presentation**
     - `templates/` renders upload forms, comparison dashboards, mismatch highlights, and source artifacts.
     - `static/` hosts Bootstrap-based styling and interactivity.

  ---

  ## Running Tests

  ```bash
  pytest
  ```

  Notable suites:
  - `test_app.py` — Flask routes and API smoke tests.
  - `test_integration.py`, `test_fixed_energy.py` — extraction/comparison regression coverage.

  ---

  ## Directory Layout

  ```
  app.py                         # Flask entry point and comparison logic
  document_extractor.py          # DOCX parsing helpers
  doc_extractor_utils.py         # Recommendation & AR utilities
  excel_keyinfo_extractor.py     # Excel parsing helpers
  templates/                     # HTML templates
  static/                        # CSS/JS assets
  agents/                        # LLM-based summary checker tooling
  tests/                         # Pytest suites (e.g. test_app.py, test_integration.py)
  uploads/                       # Temporary upload storage
  docs/, notebooks/              # Support material and experiments
  ```

  ---

  ## Function Reference

  ### `app.py`
  | Function | Purpose |
  | --- | --- |
  | `allowed_file(filename)` | Validate uploaded file extensions (`.docx`, `.xlsx`). |
  | `compare_values(doc_value, excel_value, tolerance=0.01)` | Compare two values with string normalization
  and numeric tolerance; returns mismatch metadata. |
  | `compare_general_info(doc_info, excel_info, excel_energy_data=None)` | Field-level comparison of general
  info, handling calculated fields (e.g., value per finished product). |
  | `compare_energy_data(doc_energy, excel_energy)` | Compare energy usage/costs per source and aggregate
  totals. |
  | `index()` | Render the upload form (`GET /`). |
  | `upload_files()` | Handle web uploads, run extractors, compare results, render comparison view. |
  | `api_compare()` | REST endpoint mirroring upload flow but returning JSON. |

  ### `document_extractor.py`
  | Function | Purpose |
  | --- | --- |
  | `extract_itac_report(docx_path, output="html", save_files=True)` | Central DOCX parser returning sectioned
  HTML or JSON. |
  | `build_outputs(blocks, output)` | Assemble targeted sections (general info, energy usage, ARs) from DOCX
  blocks. |
  | `extract_general_info_fields(general_info_html)` | Parse general info table into normalized numeric/string
  fields. |
  | `extract_energy_usage(annual_energy_html)` | Extract energy usage/cost data with unit metadata. |
  | Section helpers (`iter_block_items`, `extract_section_by_title`, `find_table_by_caption`, `extract_ars`,
  `extract_ar_summaries`) | Navigate DOCX structure to locate relevant sections/tables. |
  | HTML/JSON renderers (`paragraph_to_html/json`, `table_to_html/json`, `blocks_to_html/json`) | Convert DOCX
  content into serializable formats. |
  | `write_artifacts(payload, output)` | Persist extraction results to disk for inspection (HTML or JSON). |

  ### `doc_extractor_utils.py`
  | Function | Purpose |
  | --- | --- |
  | `get_recommended_summary_table_json(html)` | Convert AR summary table HTML to structured rows + totals. |
  | `get_single_ar_summary_table(ar_html)` | Extract “Savings Summary” from individual AR sections. |
  | `validate_recommendation_totals(summary_json, tolerance=0.01)` | Verify totals row matches per-AR sums. |
  | `compare_ar_with_summary(ar_data, summary_recommendations)` | Cross-check individual AR metrics against
  summary table entries. |
  | `parse_ar_summaries(ar_summary_html)` | Parse narrative AR summaries into structured list. |

  ### `excel_keyinfo_extractor.py`
  | Function | Purpose |
  | --- | --- |
  | `extract_all_structured_info(xlsx_path)` | Convenience wrapper returning general, energy/waste, and
  recommendation dictionaries. |
  | `extract_general_info_dict(xlsx_path)` | Standardize “General Info” sheet fields (SIC, NAICS, production,
  etc.). |
  | `extract_energy_waste_info_dict(xlsx_path)` | Normalize energy/waste consumption and cost data, including
  totals. |
  | `extract_recommendation_info_dict(xlsx_path)` | Parse recommendation table (ARC codes, savings, costs,
  payback). |
  | `detect_key_value_pairs`, `detect_tables`, `export_tables_to_csv` | Generic sheet scanning helpers for
  exploratory exports. |
  | CLI entrypoint (`main()`) | Scriptable interface: write JSON summary and optional CSV tables. |

  ### `agents/summary_agent.py`
  | Function | Purpose |
  | --- | --- |
  | `validate_ar_summary(ar_summary, ar_number, comparison_data)` | Compile context highlighting matches/
  differences between text and numeric data. |
  | `create_summary_checker_agent(api_key=None)` | Build a Google ADK LLM agent wired to
  `validate_ar_summary`. |
  | `check_all_ar_summaries(ar_summaries, summary_recommendations, ar_data_list, api_key=None)` | Batch-
  validate AR summaries using structured comparison data. |
  | `analyze_with_llm(validation_results, api_key=None)` | Invoke Gemini to produce narrative analysis of
  validation findings. |

  ---

  ## Data & Configuration Notes

  - **Numeric tolerance**: `compare_values` defaults to 1% relative difference. Adjust per domain needs.
  - **File size limits**: Uploads capped at 16 MB each. Modify `app.config['MAX_CONTENT_LENGTH']` if
  necessary.
  - **Temporary files**: Uploaded files are saved in `uploads/` and deleted after processing; API route uses
  `tempfile.NamedTemporaryFile`.
  - **Logging**: `logging.basicConfig(level=logging.INFO)` plus `icecream` debug output aid troubleshooting.
  - **LLM agent**: Requires `GOOGLE_API_KEY` (Gemini) when using the optional summary checker.

  ---