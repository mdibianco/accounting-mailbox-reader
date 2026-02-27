# Project Quick Start Summary

## What Was Created

A Python project for reading, classifying, and analyzing emails from Planted's accounting mailbox. The system extracts email content and attachment text, classifies emails using AI (Gemini/OpenAI), and outputs structured JSON with classification results (category, priority, confidence, extracted entities).

## Project Location

```
C:\Users\MatthiasDiBianco\accounting-mailbox-reader\
```

## Key Deliverables

### CLI Application
- **Main Entry Point**: `main.py`
- **Framework**: Python Click CLI
- **Commands**:
  - `read` - Read and analyze emails (with optional `--classify` for AI classification)
  - `preview` - Quick 5-email preview
  - `test-classify` - Test classification on a single email
  - `sync-categories` - Sync categories from Confluence
  - `config-show` - Display current configuration
  - `init` - Initialize environment

### Core Modules (src/)
1. **config.py** - Configuration management from YAML + environment variables
2. **graph_client.py** - Microsoft Graph API integration (email fetch, SharePoint upload)
3. **email_reader.py** - Email parsing, attachment discovery, classification storage
4. **attachment_analyzer.py** - Content extraction from PDF (pdfplumber/pypdf), Excel, CSV
5. **email_classifier.py** - AI classification using Gemini (default) or OpenAI
6. **gemini_cli_auth.py** - Gemini API authentication (API key or CLI OAuth)
7. **confluence_sync.py** - Category table sync from Confluence
8. **output_formatter.py** - JSON, table, and detailed text output formatting

### Classification System
- **Base prompt**: `config/classification_prompt.txt` - classification rules, priority assessment, output schema
- **Categories**: `data/categories_cache.json` - 6 categories (VEN-INV, VEN-REM, VEN-REMIT, CUST-REM-FOLLOWUP, CUST-REMIT, OTHER)
- **LLM Providers**: Gemini 2.5 Flash (default) or OpenAI GPT-4o-mini
- **Output**: Category, priority (PRIO_HIGHEST..LOW), confidence (HIGH/MEDIUM/LOW), extracted entities, reasoning

## Project Structure

```
accounting-mailbox-reader/
├── src/
│   ├── config.py                # Configuration management
│   ├── graph_client.py          # Graph API client
│   ├── email_reader.py          # Email reader + classification storage
│   ├── attachment_analyzer.py   # PDF/Excel/CSV content extraction
│   ├── email_classifier.py      # AI classification (Gemini/OpenAI)
│   ├── gemini_cli_auth.py       # Gemini authentication
│   ├── confluence_sync.py       # Confluence category sync
│   └── output_formatter.py      # Output formatting
├── config/
│   ├── settings.yaml            # Application settings
│   └── classification_prompt.txt # LLM classification prompt
├── data/
│   └── categories_cache.json    # Cached categories
├── main.py                      # CLI entry point
├── requirements.txt             # Python dependencies
└── .env                         # Credentials (not in git)
```

## Data Flow

```
CLI (main.py)
  → EmailReader
    → GraphAPIClient (fetch emails from Microsoft Graph)
    → AttachmentAnalyzer (extract text from PDF/Excel/CSV)
  → EmailClassifier (optional, with --classify)
    → Gemini or OpenAI API
    → Categories from cache
  → OutputFormatter (JSON with classification embedded)
```

## Quick Command Reference

```bash
# Activate venv
.\.venv\Scripts\activate

# Read emails with classification
python main.py read --classify --format json --output classified.json

# Test classification on most recent email
python main.py test-classify

# Sync categories from Confluence
python main.py sync-categories

# Quick preview
python main.py preview
```

## Environment Variables (.env)

```env
# Azure AD (required)
AZURE_CLIENT_ID=...
AZURE_TENANT_ID=...
AZURE_CLIENT_SECRET=...

# LLM (at least one)
LLM_PROVIDER=gemini
GEMINI_API_KEY=...
# OPENAI_API_KEY=...

# Optional
CONFLUENCE_EMAIL=...
CONFLUENCE_API_TOKEN=...
LOCAL_FOLDER_PATH=C:\Users\MatthiasDiBianco\emails\emails
```

## Roadmap

- **Phase 0**: ✅ Email reading & attachment extraction
- **Phase 1**: ✅ AI classification (categories, priority, entity extraction)
- **Phase 1.5**: Category-specific deep extraction (Pass 2 prompts, reminder levels)
- **Phase 2**: Business Central matching (vendor/invoice lookup)
- **Phase 3**: Autonomous processing workflows
