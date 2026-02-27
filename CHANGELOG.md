# Changelog

All notable changes to this project will be documented in this file.

## [0.1.0] - 2026-02-10

### Foundation Phase Complete ✓

#### Added

**Core Features:**
- ✅ Microsoft Graph API client for reading shared mailboxes
- ✅ Email reader with full metadata extraction (from, subject, body, attachments)
- ✅ Attachment analyzer with multi-format support
  - PDF extraction (pdfplumber + pypdf2 fallback)
  - Excel workbook parsing (openpyxl)
  - CSV file reading
  - Image placeholder (OCR planned for Phase 2)
- ✅ Output formatters
  - JSON format (structured data export)
  - Console table format (quick overview)
  - Detailed text format (full content with extracted attachments)

**CLI Commands:**
- `main.py read` - Read and analyze emails with flexible options
  - `--days` - Filter by date range
  - `--max` - Limit email count
  - `--format` - Choose output format (json, table, detailed)
  - `--output` - Save to file
  - `--no-attachments` - Skip attachment processing
  - `--no-body` - Skip full body content
  - `--search` - OData query filtering
- `main.py preview` - Quick preview mode (5 most recent emails)
- `main.py config-show` - Display configuration
- `main.py init` - Initialize `.env` file

**Configuration:**
- YAML-based settings (`config/settings.yaml`)
- Environment variable support (.env file)
- Type-hinted dataclasses for all data models:
  - `Email` - Email message with full metadata
  - `Attachment` - Email attachment metadata
  - `ExtractedText` - Attachment content extraction result

**Helper Scripts:**
- `run.bat` - Windows convenience script
- `run.sh` - macOS/Linux convenience script

**Documentation:**
- `README.md` - Complete feature overview and usage guide
- `SETUP.md` - Comprehensive setup instructions with Azure AD walkthrough
- `EXAMPLES.md` - Practical usage examples and automation patterns
- `.github/copilot-instructions.md` - Development guidelines

**Project Structure:**
```
accounting-mailbox-reader/
├── src/
│   ├── config.py              # Configuration management
│   ├── graph_client.py        # Microsoft Graph API wrapper
│   ├── email_reader.py        # Email reading & parsing
│   ├── attachment_analyzer.py # Document content extraction
│   └── output_formatter.py    # Output formatting (JSON, console, detailed)
├── config/
│   └── settings.yaml          # Application configuration
├── main.py                    # Click CLI entry point
├── requirements.txt           # Python dependencies
├── .env.example               # Environment template
└── [documentation files]
```

#### Security Features
- ✅ Read-only mailbox access (no email mutations)
- ✅ Environment variable-based credential management
- ✅ `.gitignore` protection for `.env` file
- ✅ Comprehensive audit logging ready
- ✅ Azure AD OAuth 2.0 integration

#### Testing
- ✅ CLI tested: `config-show`, `init`, `preview`
- ✅ Python environment configured and dependencies installed
- ✅ All modules structurally complete with type hints and docstrings

### Dependencies Added
```
python-dotenv==1.0.0          # Environment variable management
requests==2.31.0              # HTTP client
msgraph-core==0.2.2           # Graph API SDK
azure-identity==1.14.0        # Azure authentication
pydantic==2.5.0               # Data validation
pyyaml==6.0.1                 # YAML configuration
pypdf2==4.0.1                 # PDF text extraction (fallback)
pdfplumber==0.10.3            # Advanced PDF parsing
openpyxl==3.11.0              # Excel file reading
python-dateutil==2.8.2        # Date utilities
Pillow==10.1.0                # Image processing (OCR support)
click==8.1.7                  # CLI framework
tabulate==0.9.0               # Console table formatting
```

### Known Limitations
- OCR for images not yet implemented (placeholder only)
- HTML-to-text conversion uses regex (could be improved with BeautifulSoup)
- Single-threaded email processing (async planned for Phase 3+)
- No classification or risk assessment yet (Phase 1+)
- No Business Central integration yet (Phase 3+)

### Next Phase Items (Phase 1+)

**Phase 1: AI Classification**
- [ ] Implement `AccountingClassifier` with Claude API
- [ ] Add accounting-specific email categories
- [ ] Create risk assessment engine
- [ ] Jira ticket creation support

**Phase 2: Autonomous Processing**
- [ ] Implement payment reminder workflow
- [ ] Implement remittance advice matching
- [ ] Add draft response generation

**Phase 3: Business Central**
- [ ] Implement `BusinessCentralClient`
- [ ] Query invoice and vendor data
- [ ] Add payment status checking
- [ ] Autonomous processing with BC context

**Future Enhancements**
- [ ] OCR support for scanned documents
- [ ] Async/concurrent email processing
- [ ] Webhook support for real-time processing
- [ ] Multi-language support (Français, Deutsch)
- [ ] Jira workflow automation integration

## Installation & Usage

See [SETUP.md](SETUP.md) for complete setup instructions.

```bash
# Quick start
python main.py init
python main.py config-show
python main.py preview
```

## Architecture

See [README.md](README.md) for complete architecture documentation.

## Development

This project follows the development guidelines in `.github/copilot-instructions.md`

### Code Standards
- Type hints on all functions
- Docstrings for classes and public methods
- Dataclasses for data models
- Single responsibility principle
- Comprehensive logging

### Running

```bash
.venv\Scripts\python main.py --help
```

## Contributors

Initial foundation (Phase 0): Implemented by GitHub Copilot
