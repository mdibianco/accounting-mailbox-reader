# Accounting Mailbox Reader

A Python CLI tool for reading, classifying, and analyzing emails from Planted's accounting mailbox (`accounting@eatplanted.com`). Integrates with Microsoft Graph API to fetch emails, extracts content from attachments, and classifies emails using AI (Gemini or OpenAI).

**Current Version**: 0.2.0 (Phase 1 - Email Classification)

## Features

### Email Reading
- ✅ Read emails from shared mailbox via Microsoft Graph API
- ✅ Extract full email body and metadata (sender, date, priority)
- ✅ Automatic attachment content extraction:
  - **PDF**: Text extraction via pdfplumber or pypdf
  - **Excel**: Cell content extraction
  - **CSV**: Direct file content
  - **Images**: Placeholder for future OCR support
- ✅ Full attachment text preserved in JSON output

### AI Classification (Phase 1)
- ✅ Classify emails into accounting categories (VEN-INV, VEN-REM, CUST-REMIT, etc.)
- ✅ Context-aware priority assessment (PRIO_HIGHEST to PRIO_LOW)
- ✅ 3-level confidence system (HIGH / MEDIUM / LOW)
- ✅ Entity extraction (vendor, invoice number, amount, due date)
- ✅ Full attachment text sent to LLM for accurate classification
- ✅ Multiple LLM providers: Gemini (default) or OpenAI
- ✅ Categories synced from Confluence page

### Output & Storage
- ✅ Multiple output formats: Table, JSON, Detailed text
- ✅ Classification results embedded in JSON output
- ✅ Save to local folder, SharePoint, or via Power Automate
- ✅ Flexible filtering (date range, max count, OData search)

## Quick Start

### 1. Prerequisites

- Python 3.8+
- Microsoft Azure AD application registration (for Graph API access)
- Internet connection

### 2. Installation

```bash
# Clone or download the repository
cd accounting-mailbox-reader

# Create and activate virtual environment
python -m venv .venv
.\.venv\Scripts\activate  # Windows
# source venv/bin/activate  # macOS/Linux

# Install dependencies
pip install -r requirements.txt

# Initialize .env file
python main.py init
```

### 3. Azure Setup

You need an Azure AD application with permission to read the shared mailbox:

**Steps:**
1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory > App registrations > New registration**
3. Name: `Accounting Mailbox Reader`
4. **Add API permissions:**
   - Microsoft Graph → Delegated permissions
   - Add: `Mail.Read.Shared`, `Mail.ReadWrite.Shared`
5. **Create a client secret:**
   - Certificates & secrets → New client secret
   - Copy the value immediately (you won't see it again)
6. Get:
   - Application (client) ID
   - Directory (tenant) ID
   - Client secret value

**Fill in your .env file:**
```env
AZURE_CLIENT_ID=your-app-client-id
AZURE_TENANT_ID=your-tenant-id
AZURE_CLIENT_SECRET=your-client-secret
ACCOUNTING_MAILBOX=accounting@eatplanted.com
```

### 4. First Run

```bash
# Verify configuration
python main.py config-show

# Preview most recent 5 emails (1 day back)
python main.py preview

# Read full mailbox (last 7 days, up to 50 emails)
python main.py read
```

## Usage

### Commands

#### `read` - Read and analyze emails

```bash
# Basic usage - table format, last 7 days, max 50 emails
python main.py read

# With AI classification
python main.py read --classify --format json --output classified.json

# Read last 14 days, max 100 emails
python main.py read --days 14 --max 100

# JSON output (best for programmatic use)
python main.py read --format json

# Save to local folder and upload to SharePoint
python main.py read --upload-sharepoint

# Skip attachment extraction (faster, but no PDF text for classification)
python main.py read --no-attachments

# Search for specific emails (OData syntax)
python main.py read --search "from:vendor@example.com"
```

#### `test-classify` - Test classification on a single email

```bash
# Classify the most recent email (useful for testing)
python main.py test-classify
```

#### `sync-categories` - Sync categories from Confluence

```bash
# Fetch latest category definitions from Confluence
python main.py sync-categories
```

#### `preview` - Quick preview of recent emails

```bash
python main.py preview
```

#### `config-show` / `init`

```bash
python main.py config-show   # Display current configuration
python main.py init           # Initialize .env file
```

## Output Formats

### Table Format (Default)
```
════════════════════════════════════════════════════════════════════════════════
From                Subject                                       Received  Attachments
════════════════════════════════════════════════════════════════════════════════
vendor@example.com  INV-2026-001 Payment Reminder                 2026-02-10  reminder.pdf
customer@abc.com    Remittance Advice - Payment Sent              2026-02-09  None
bank@ubs.com        Monthly Statement February 2026               2026-02-08  statement.pdf
════════════════════════════════════════════════════════════════════════════════
```

### JSON Format (with `--classify`)
```json
{
  "timestamp": "2026-02-10T14:30:00.000000",
  "count": 3,
  "emails": [
    {
      "id": "message-id-123",
      "from": {
        "email": "vendor@example.com",
        "name": "Acme Corp"
      },
      "subject": "INV-2026-001 Payment Reminder",
      "received_datetime": "2026-02-10T10:25:00Z",
      "body_preview": "Dear Planted, we have not received...",
      "body": "Full email body here...",
      "has_attachments": true,
      "is_read": false,
      "importance": "high",
      "attachments": [
        {
          "id": "att-123",
          "name": "reminder.pdf",
          "content_type": "application/pdf",
          "size": 245000,
          "extracted_text": {
            "filename": "reminder.pdf",
            "content_type": "application/pdf",
            "text": "Full extracted PDF text here...",
            "extraction_method": "pdf_pdfplumber",
            "success": true
          }
        }
      ],
      "classification": {
        "confidence_level": "HIGH",
        "primary_category": {
          "id": "VEN-REM",
          "name": "Vendor Payment Reminder"
        },
        "secondary_category": null,
        "priority": "PRIO_HIGH",
        "requires_manual_review": false,
        "extracted_entities": {
          "vendor": "Acme Corp",
          "invoice_number": "INV-2026-001",
          "amount": 5430.00,
          "currency": "CHF",
          "due_date": "2026-01-25"
        },
        "reasoning": "Email from external vendor references unpaid invoice..."
      }
    }
  ]
}
```

### Detailed Format
Includes full email body and extracted attachment text with headers for easy navigation.

## Configuration

### settings.yaml

Located in `config/settings.yaml`:

```yaml
accounting_triage:
  mailbox: "accounting@eatplanted.com"
  dry_run: true
  email_reader:
    max_emails: 50       # Default max emails
    days_back: 7         # Default days to read back
  attachments:
    enabled: true
    supported_formats:
      - ".pdf"
      - ".xlsx"
      - ".xls"
      - ".csv"
      - ".png"
      - ".jpg"
    max_size_mb: 25      # Skip attachments larger than this
```

### Environment Variables (.env)

**Required:**
- `AZURE_CLIENT_ID` - Azure AD application client ID
- `AZURE_TENANT_ID` - Azure AD tenant ID
- `AZURE_CLIENT_SECRET` - Azure AD application secret

**Classification (at least one LLM provider):**
- `LLM_PROVIDER` - `gemini` (default) or `openai`
- `GEMINI_API_KEY` - Google AI Studio API key (for Gemini)
- `OPENAI_API_KEY` - OpenAI API key (if using OpenAI)

**Optional:**
- `CONFLUENCE_EMAIL` - Confluence email for category sync
- `CONFLUENCE_API_TOKEN` - Confluence API token
- `LOCAL_FOLDER_PATH` - Local folder for email JSON output
- `POWER_AUTOMATE_FLOW_URL` - Power Automate HTTP trigger URL

## Project Structure

```
accounting-mailbox-reader/
├── src/
│   ├── __init__.py
│   ├── config.py              # Configuration management
│   ├── graph_client.py        # Microsoft Graph API client
│   ├── email_reader.py        # Email reading & parsing
│   ├── attachment_analyzer.py # Attachment content extraction
│   ├── email_classifier.py   # AI classification (Gemini/OpenAI)
│   ├── gemini_cli_auth.py    # Gemini API authentication
│   ├── confluence_sync.py    # Category sync from Confluence
│   └── output_formatter.py    # Output formatting (JSON, console)
├── config/
│   ├── settings.yaml          # Application settings
│   └── classification_prompt.txt  # Base LLM classification prompt
├── data/
│   └── categories_cache.json  # Cached categories from Confluence
├── main.py                    # CLI entry point
├── requirements.txt           # Python dependencies
└── .env.example               # Environment variable template
```

## Architecture

### Component Overview

```
┌──────────────────────────────────────────────────────────────┐
│  CLI (main.py with Click)                                    │
│  read | preview | test-classify | sync-categories | ...      │
└─────────┬──────────────────────────────┬─────────────────────┘
          │                              │
┌─────────▼──────────────────┐  ┌────────▼──────────────────────┐
│ EmailReader                │  │ EmailClassifier               │
│  - Fetch emails            │  │  - Gemini or OpenAI           │
│  - Download attachments    │  │  - Category + Priority        │
│  - Extract text (PDF/XLS)  │  │  - Confidence + Entities      │
└──┬──────────────┬──────────┘  └────────┬──────────────────────┘
   │              │                      │
┌──▼───────┐ ┌───▼───────────────┐ ┌────▼────────────────────────┐
│ Graph    │ │ Attachment        │ │ Categories                   │
│ Client   │ │ Analyzer          │ │ (Confluence sync or cache)   │
└──────────┘ └───────────────────┘ └──────────────────────────────┘

OutputFormatter: JSON (with classification) | Console table | Detailed text
```

## Categories

Categories are defined in Confluence and synced to `data/categories_cache.json`:

| ID | Name | Description |
|----|------|-------------|
| VEN-INV | Vendor Invoice | Original invoice from vendor (not a reminder) |
| VEN-REM | Vendor Payment Reminder | Payment reminder from vendor to Planted |
| VEN-REMIT | Vendor Remittance Advice | Vendor requests payment details from Planted |
| CUST-REM-FOLLOWUP | Customer Follow-up | Customer responds to a reminder Planted sent |
| CUST-REMIT | Customer Remittance Advice | Customer informs Planted of payments made |
| OTHER | Uncategorized | Requires manual review |

## Limitations & Known Issues

1. **OCR not yet implemented** - Image files show placeholder text
2. **HTML in email bodies** - Converted to plain text via regex (could use BeautifulSoup for better results)
3. **Large attachments** - Skipped if > 25MB (configurable)
4. **Read-only access** - Only reads emails, doesn't compose or move messages (by design)

## Roadmap

- **Phase 0**: ✅ Email reading & attachment extraction
- **Phase 1**: ✅ AI classification (categories, priority, entity extraction)
- **Phase 1.5**: Category-specific deep extraction (Pass 2 prompts, reminder levels, structured entities)
- **Phase 2**: Microsoft Business Central matching (vendor/invoice lookup)
- **Phase 3**: Autonomous processing workflows
- **Phase 4**: Jira ticket creation and workflow management

## Troubleshooting

### No emails found
- Check mailbox address in `.env`
- Verify Azure credentials are correct
- Check date range (`--days`) - may not have emails in that period
- Ensure delegated permissions are granted in Azure

### Attachment extraction fails
- Check file format is in supported list
- Verify attachment size < 25MB (or adjust in config)
- Some PDF files may not extract well (encrypted, scanned images)

### Azure authentication fails
- Double-check credentials in `.env`
- Verify application is registered in Azure AD
- Ensure required API permissions are granted
- Try creating a new client secret (old one may be revoked)

## Development

### Running Tests (Future)
```bash
pytest tests/
```

### Logging
Set environment variable to increase verbosity:
```bash
# In your script or terminal
export LOGLEVEL=DEBUG
python main.py read
```

## Security Notes

- **Never commit `.env` file** - Add to `.gitignore` ✓
- **Client secret is sensitive** - Keep out of version control
- **Emails may contain PII** - Be careful with output files
- **Azure permissions are scoped** - Read-only for safe operation

## Contributing

1. Create a feature branch: `git checkout -b feature/my-feature`
2. Make changes and test thoroughly
3. Submit a pull request

## License

Internal Planted tool

## Contact

See `.github/copilot-instructions.md` for project maintenance guidelines.
