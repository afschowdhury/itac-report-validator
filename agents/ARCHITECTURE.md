# AR Summary Checker Agent - Architecture

## System Architecture

```
┌─────────────────────────────────────────────────────────────────────┐
│                    ITAC Report Validation System                     │
└─────────────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────────────┐
│                         INPUT SOURCES                                │
├─────────────────────────────────────────────────────────────────────┤
│  📄 ar_summary.html                                                  │
│  📄 recommendation_summary_table.html                                │
│  📄 AR_01.html, AR_02.html, ... AR_N.html                           │
└─────────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────────┐
│                      DATA EXTRACTION LAYER                           │
│                    (doc_extractor_utils.py)                          │
├─────────────────────────────────────────────────────────────────────┤
│  • parse_ar_summaries()                                             │
│  • get_recommended_summary_table_json()                             │
│  • get_single_ar_summary_table()                                    │
│  • compare_ar_with_summary()                                        │
└─────────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────────┐
│                     AI AGENT LAYER (ADK)                             │
│                     (summary_agent.py)                               │
├─────────────────────────────────────────────────────────────────────┤
│                                                                      │
│  ┌────────────────────────────────────────────────────────┐         │
│  │     AR Summary Checker Agent (LlmAgent)                │         │
│  │     Model: Gemini 2.0 Flash                            │         │
│  │     Type: LLM Agent                                    │         │
│  └────────────────────────────────────────────────────────┘         │
│                           │                                          │
│                           ├─► validate_ar_summary()                 │
│                           │   (Function Tool)                        │
│                           │                                          │
│                           ├─► check_all_ar_summaries()              │
│                           │   (Batch Processor)                      │
│                           │                                          │
│                           └─► analyze_with_llm()                    │
│                               (AI Analysis)                          │
│                                                                      │
└─────────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────────┐
│                        OUTPUT LAYER                                  │
├─────────────────────────────────────────────────────────────────────┤
│  📊 ar_summary_validation_results.json                              │
│  📝 ar_summary_analysis.txt                                         │
│  📈 Validation Reports                                              │
└─────────────────────────────────────────────────────────────────────┘
```

## Google ADK Component Architecture

```
┌─────────────────────────────────────────────────────────────────────┐
│                      GOOGLE ADK FRAMEWORK                            │
└─────────────────────────────────────────────────────────────────────┘

┌──────────────────┐    ┌──────────────────┐    ┌──────────────────┐
│   LLM Agents     │    │ Workflow Agents  │    │  Custom Agents   │
│                  │    │                  │    │                  │
│ • Dynamic        │    │ • Sequential     │    │ • Fully          │
│   Reasoning      │    │ • Loop           │    │   Customizable   │
│ • Tool Use       │    │ • Parallel       │    │                  │
│ • Transfers      │    │ • Deterministic  │    │                  │
│                  │    │                  │    │                  │
│  [USED HERE]     │    │                  │    │                  │
└──────────────────┘    └──────────────────┘    └──────────────────┘
        │
        ↓
┌─────────────────────────────────────────────────────────────────────┐
│                         TOOLS ECOSYSTEM                              │
├─────────────────────────────────────────────────────────────────────┤
│  • Function Tools (Custom Python functions)    [USED HERE]          │
│  • Built-in Tools (Search, Code Exec, etc.)                         │
│  • Third-party Tools (LangChain, CrewAI)                            │
│  • Google Cloud Tools (Vertex AI, etc.)                             │
│  • MCP Tools (Model Context Protocol)                               │
│  • OpenAPI Tools                                                    │
└─────────────────────────────────────────────────────────────────────┘
```

## Data Flow Diagram

```
┌─────────────────┐
│  HTML Files     │
│  (Extracted)    │
└────────┬────────┘
         │
         ↓
┌────────────────────────────────────────────────────────────────┐
│  Step 1: Parse AR Summaries                                    │
│  ─────────────────────────────────────────────────────────────│
│  parse_ar_summaries(ar_summary_html)                          │
│  Output: [                                                     │
│    { ar_no: 1, ar_summary: "Text..." },                       │
│    { ar_no: 2, ar_summary: "Text..." },                       │
│    ...                                                         │
│  ]                                                             │
└────────┬───────────────────────────────────────────────────────┘
         │
         ↓
┌────────────────────────────────────────────────────────────────┐
│  Step 2: Parse Recommendation Summary Table                    │
│  ─────────────────────────────────────────────────────────────│
│  get_recommended_summary_table_json(rec_summary_html)         │
│  Output: {                                                     │
│    recommendations: [                                          │
│      { ar_number: 1, electricity_savings: 100000, ... },      │
│      { ar_number: 2, electricity_savings: 50000, ... },       │
│      ...                                                       │
│    ]                                                           │
│  }                                                             │
└────────┬───────────────────────────────────────────────────────┘
         │
         ↓
┌────────────────────────────────────────────────────────────────┐
│  Step 3: Parse Individual AR Data                              │
│  ─────────────────────────────────────────────────────────────│
│  For each AR HTML:                                             │
│    get_single_ar_summary_table(ar_html)                       │
│  Output: [                                                     │
│    { ar_number: 1, data: { electricity_savings: ..., } },     │
│    { ar_number: 2, data: { electricity_savings: ..., } },     │
│    ...                                                         │
│  ]                                                             │
└────────┬───────────────────────────────────────────────────────┘
         │
         ↓
┌────────────────────────────────────────────────────────────────┐
│  Step 4: Validate Each AR                                      │
│  ─────────────────────────────────────────────────────────────│
│  For each AR:                                                  │
│    1. Compare AR data with summary table                       │
│       → compare_ar_with_summary(ar_data, summary_recs)        │
│    2. Validate summary text against numbers                    │
│       → validate_ar_summary(summary, ar_no, comparison)       │
│  Output: {                                                     │
│    ar_number: 1,                                              │
│    validation: { has_differences: false, ... },               │
│    comparison: { matches: [...], differences: [...] }         │
│  }                                                             │
└────────┬───────────────────────────────────────────────────────┘
         │
         ↓
┌────────────────────────────────────────────────────────────────┐
│  Step 5: AI-Powered Analysis (Optional)                        │
│  ─────────────────────────────────────────────────────────────│
│  analyze_with_llm(validation_results, api_key)                │
│                                                                │
│  Gemini analyzes all results to:                              │
│  • Identify ARs with critical issues                          │
│  • Find common patterns                                       │
│  • Assess severity                                            │
│  • Provide recommendations                                    │
│                                                                │
│  Output: Comprehensive text analysis report                    │
└────────┬───────────────────────────────────────────────────────┘
         │
         ↓
┌────────────────────────────────────────────────────────────────┐
│  Final Output: Validation Reports                              │
│  ─────────────────────────────────────────────────────────────│
│  • JSON: Structured validation results                         │
│  • Text: AI-generated analysis report                          │
└────────────────────────────────────────────────────────────────┘
```

## Agent Decision Flow

```
┌─────────────────────────────────────────────────────────────┐
│  AR Summary Checker Agent                                   │
│  (LlmAgent with Gemini 2.0 Flash)                          │
└─────────────────────────────────────────────────────────────┘
                        │
                        ↓
         ┌──────────────────────────────┐
         │  Agent Receives Input:       │
         │  • AR summary text           │
         │  • Numerical data (matches)  │
         │  • Discrepancies (diffs)     │
         └──────────────┬───────────────┘
                        │
                        ↓
         ┌──────────────────────────────┐
         │  Agent Analyzes:             │
         │  1. Read summary text        │
         │  2. Compare with numbers     │
         │  3. Identify inconsistencies │
         │  4. Assess severity          │
         └──────────────┬───────────────┘
                        │
                        ↓
         ┌──────────────────────────────┐
         │  Agent Uses Tools:           │
         │  • validate_ar_summary()     │
         │    (Function Tool)           │
         └──────────────┬───────────────┘
                        │
                        ↓
         ┌──────────────────────────────┐
         │  Agent Generates Output:     │
         │  • Validation results        │
         │  • Issue descriptions        │
         │  • Severity ratings          │
         │  • Recommendations           │
         └──────────────────────────────┘
```

## Technology Stack

```
┌─────────────────────────────────────────────────────────────┐
│  Layer                │  Technology                         │
├───────────────────────┼─────────────────────────────────────┤
│  AI Framework         │  Google ADK (Agent Development Kit) │
│  LLM Model            │  Gemini 2.0 Flash                   │
│  Language             │  Python 3.x                         │
│  HTML Parsing         │  BeautifulSoup4                     │
│  Data Processing      │  Pandas                             │
│  Document Handling    │  python-docx, openpyxl              │
│  Web Framework        │  Flask (existing)                   │
└─────────────────────────────────────────────────────────────┘
```

## Deployment Options (via ADK)

```
┌─────────────────────────────────────────────────────────────┐
│  1. Local Development                                       │
│     • Run directly with Python                              │
│     • Interactive notebooks                                 │
│     • Demo scripts                                          │
├─────────────────────────────────────────────────────────────┤
│  2. Vertex AI Agent Engine                                  │
│     • Managed deployment                                    │
│     • Scalable infrastructure                               │
│     • Built-in monitoring                                   │
├─────────────────────────────────────────────────────────────┤
│  3. Cloud Run                                               │
│     • Containerized deployment                              │
│     • Serverless scaling                                    │
│     • HTTP/gRPC endpoints                                   │
├─────────────────────────────────────────────────────────────┤
│  4. GKE (Google Kubernetes Engine)                          │
│     • Full control                                          │
│     • Custom orchestration                                  │
│     • Enterprise-grade                                      │
├─────────────────────────────────────────────────────────────┤
│  5. Custom Infrastructure                                   │
│     • Docker containers                                     │
│     • Any cloud/on-prem                                     │
│     • Complete flexibility                                  │
└─────────────────────────────────────────────────────────────┘
```

## Extension Points

```
┌─────────────────────────────────────────────────────────────┐
│  Future Multi-Agent Architecture                            │
└─────────────────────────────────────────────────────────────┘

           ┌──────────────────────────┐
           │  Master Validator Agent   │
           │  (Sequential Workflow)    │
           └───────────┬──────────────┘
                       │
       ┌───────────────┼───────────────┐
       ↓               ↓               ↓
┌─────────────┐ ┌─────────────┐ ┌─────────────┐
│  Summary    │ │  Energy     │ │  Financial  │
│  Checker    │ │  Validator  │ │  Validator  │
│  Agent      │ │  Agent      │ │  Agent      │
└─────────────┘ └─────────────┘ └─────────────┘
  [CURRENT]         [FUTURE]        [FUTURE]
```

