# Complete Setup Guide for Accounting Mailbox Reader

## System Requirements

- **OS**: Windows, macOS, or Linux
- **Python**: 3.8 or higher
- **RAM**: Minimum 512MB free
- **Disk Space**: ~500MB (including virtual environment)
- **Internet**: Required for Microsoft Graph API

## Step 1: Clone or Download the Project

```bash
# Option A: If you have Git
git clone <repository-url>
cd accounting-mailbox-reader

# Option B: Download ZIP and extract
cd accounting-mailbox-reader
```

## Step 2: Create Virtual Environment

### Windows (PowerShell)
```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

### Windows (Command Prompt)
```cmd
python -m venv .venv
.venv\Scripts\activate.bat
```

### macOS/Linux
```bash
python3 -m venv venv
source venv/bin/activate
```

**Verify activation:**
```bash
# You should see (.venv) or (venv) at the start of your terminal prompt
```

## Step 3: Install Dependencies

```bash
pip install -r requirements.txt
```

**Verify installation:**
```bash
pip list
# Should show: click, pydantic, pyyaml, requests, azure-identity, etc.
```

## Step 4: Configure Azure Active Directory

### 4.1 Create an Azure AD Application

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** → **App registrations**
3. Click **New registration**
   - **Name**: `Accounting Mailbox Reader`
   - **Supported account types**: Select "Accounts in this organizational directory only"
   - Click **Register**

### 4.2 Grant API Permissions

1. In your app, go to **API permissions**
2. Click **Add a permission**
3. Select **Microsoft Graph**
4. Select **Delegated permissions**
5. Search for and add these permissions:
   - `Mail.Read.Shared` - Read emails from shared mailbox
   - `Mail.ReadWrite.Shared` - Move/mark emails as read
   
   *(For full functionality, also add:)*
   - `Mail.Send.Shared` - Send responses (future phase)

6. Click **Grant admin consent** (requires admin privileges)

### 4.3 Create a Client Secret

1. Go to **Certificates & secrets**
2. Under "Client secrets", click **New client secret**
3. Set expiration to "24 months" or as your security policy dictates
4. **Copy the VALUE immediately** (you won't be able to see it again)

### 4.4 Collect Required Information

On your app's **Overview** page, copy these values:
- **Application (client) ID**
- **Directory (tenant) ID**

## Step 5: Configure .env File

### 5.1 Initialize .env from template

```bash
# Windows (from .venv activated terminal)
python main.py init

# Or manually copy
copy .env.example .env
```

### 5.2 Fill in Your Credentials

Edit `.env` and replace the placeholder values:

```env
# Azure Active Directory Configuration
AZURE_CLIENT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
AZURE_TENANT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
AZURE_CLIENT_SECRET=your-very-long-secret-value-here

# Mailbox Configuration
ACCOUNTING_MAILBOX=accounting@eatplanted.com
DRY_RUN=true
```

**⚠️ SECURITY WARNING:**
- Never commit `.env` to Git (already in `.gitignore`)
- Never share your client secret
- Treat this file like a password

## Step 6: Verify Configuration

```bash
python main.py config-show
```

Expected output:
```
============================================================
CURRENT CONFIGURATION
============================================================

Mailbox:           accounting@eatplanted.com
Dry Run Mode:      True
Max Emails:        50
Days Back:         7
Max Attachment Size: 25 MB
Attachment Formats: .pdf, .xlsx, .xls, .csv, .png, .jpg

Azure Setup:
  Client ID:     ✓ Configured
  Tenant ID:     ✓ Configured
  Client Secret: ✓ Configured

============================================================
```

## Step 7: Test Connection

### 7.1 Preview Recent Emails

```bash
python main.py preview
```

This will attempt to read the 5 most recent emails from the last 24 hours.

**Success output:**
```
✓ Preview of 3 most recent emails

┌────────────────────┬──────────────────────────────┬───────────┬──────┬──────────────┐
│ From               │ Subject                      │ Received  │ Imp. │ Attachments  │
├────────────────────┼──────────────────────────────┼───────────┼──────┼──────────────┤
│ vendor@example.com │ Invoice INV-2026-001         │ 2026-02-10│ high │ invoice.pdf  │
│ customer@abc.com   │ Payment Received             │ 2026-02-09│ norm │ None         │
│ bank@ubs.com       │ Monthly Statement            │ 2026-02-08│ norm │ statement.pdf│
└────────────────────┴──────────────────────────────┴───────────┴──────┴──────────────┘
```

**Error troubleshooting:**

| Error | Solution |
|-------|----------|
| `AADSTS900023: Client error` | Check AZURE_CLIENT_SECRET value (typo or expired) |
| `Failed to get access token` | Check AZURE_TENANT_ID and AZURE_CLIENT_ID |
| `No emails found` | Check ACCOUNTING_MAILBOX address, or try `--days 30` |
| `ModuleNotFoundError: click` | Ensure virtual environment is activated |

## Step 8: Read Full Mailbox

```bash
# Read last 7 days, display as table (default)
python main.py read

# Read last 14 days, max 100 emails, JSON format
python main.py read --days 14 --max 100 --format json

# Save detailed output to file
python main.py read --format detailed --output emails.txt
```

## Step 9: Understand the Output

### Table Format (Default)
Quick overview suitable for command-line viewing.

### JSON Format
```bash
python main.py read --format json --output emails.json
```
Best for programmatic use or data import.

### Detailed Format
```bash
python main.py read --format detailed --output emails.txt
```
Complete email bodies and extracted attachment content.

## Advanced Usage

### Filtering Emails

```bash
# Read only emails from a specific sender
python main.py read --search "from:vendor@example.com"

# Read only unread emails in last 7 days
python main.py read --search "isRead:false"

# More complex OData queries
python main.py read --search "from:vendor AND subject:invoice"
```

### Performance Options

```bash
# Skip attachment extraction for speed
python main.py read --no-attachments

# Skip full body content
python main.py read --no-body

# Combine for fastest preview
python main.py preview --no-attachments
```

## Using Helper Scripts

### Windows (PowerShell/Command Prompt)

```bash
# Show configuration
run.bat config

# Preview emails
run.bat preview

# Read emails with options
run.bat read --format json --output emails.json

# Enter interactive shell with activated venv
run.bat shell
```

### macOS/Linux

```bash
chmod +x run.sh  # Make script executable first

# Show configuration
./run.sh config

# Preview emails
./run.sh preview

# Read emails with options
./run.sh read --format json --output emails.json

# Enter interactive shell
./run.sh shell
```

## Troubleshooting

### Virtual Environment Issues

**"python: command not found" or "python is not recognized"**
```bash
# On some systems, use python3 instead
python3 main.py config-show
```

**"ModuleNotFoundError: No module named 'X'"**
```bash
# Ensure venv is activated, then reinstall
pip install -r requirements.txt
```

### Azure Authentication Issues

**"AADSTS65001: User or admin has not consented"**
- Go back to App Registration → API permissions
- Click **Grant admin consent for [Organization]**

**"AADSTS700036: Client does not have access to the resource"**
- Verify `/Mail.Read.Shared` permission is granted
- May need to wait a few minutes after granting permissions

**"Invalid client secret"**
- Create a new client secret and update `.env`
- Old secrets may expire after 24 months

### Email Reading Issues

**"No emails found"**
- Check ACCOUNTING_MAILBOX in `.env`
- Try increasing days: `--days 30` or `--days 90`
- Verify service account has access to the mailbox

**Attachment extraction fails**
- Some encrypted PDFs cannot be extracted
- Verify file size < 25MB (or change `max_attachment_size_mb` in `config/settings.yaml`)
- Image files will show "OCR not yet implemented" (planned for Phase 2)

## Next Steps

### For Testing
1. Run `python main.py preview` regularly to verify connectivity
2. Use `--format json` to export data for analysis
3. Check `logs/` directory for detailed operation records

### For Production
1. Create a dedicated Azure service account for the mailbox reader
2. Set up scheduled task or cron job to run periodically
3. Configure email alerts for errors (see logging configuration)
4. Set `DRY_RUN=false` in `.env` when ready for mutations (future phases)

### For Development
See [README.md](README.md) for architecture and next phases.
See [.github/copilot-instructions.md](.github/copilot-instructions.md) for development guidelines.

## Support

### Getting Help
1. Check [README.md](README.md) for feature documentation
2. Review logs: `logs/accounting_triage.log`
3. Test with `--format json` to inspect raw data
4. Enable debug logging by setting `LOGLEVEL=DEBUG`

### Reporting Issues
Include:
- Python version: `python --version`
- Error message and stack trace
- Command that failed
- (Sanitized) `.env` values
