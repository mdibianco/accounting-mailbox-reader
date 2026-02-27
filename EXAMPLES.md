# Usage Examples

This document provides practical examples of how to use the Accounting Mailbox Reader in different scenarios.

## Table of Contents
1. [Basic Usage](#basic-usage)
2. [Output Formats](#output-formats)
3. [Filtering & Searching](#filtering--searching)
4. [Attachment Handling](#attachment-handling)
5. [Advanced Scenarios](#advanced-scenarios)
6. [Automation](#automation)

## Basic Usage

### Example 1: Quick Email Preview
```bash
python main.py preview
```
Shows the 5 most recent emails from the last 24 hours. Perfect for quick daily checks.

**Output:**
```
✓ Preview of 3 most recent emails

┌────────────────────┬────────────────────────┬───────────┬──────┬─────────┐
│ From               │ Subject                │ Received  │ Imp. │ Attachm.│
├────────────────────┼────────────────────────┼───────────┼──────┼─────────┤
│ vendor@abc.com     │ Invoice INV-2026-001   │ 2026-02-10│ high │ inv.pdf │
│ bank@ubs.com       │ Payment Confirmation   │ 2026-02-10│ norm │ None    │
│ customer@xyz.com   │ Remittance Advice      │ 2026-02-09│ norm │ advice. │
└────────────────────┴────────────────────────┴───────────┴──────┴─────────┘
```

### Example 2: Read Last 14 Days
```bash
python main.py read --days 14 --max 100
```
Extends the date range and increases email count limit.

### Example 3: Configuration Check
```bash
python main.py config-show
```
Verify all settings before running batch operations.

## Output Formats

### Example 4: Export to JSON
```bash
python main.py read --days 7 --format json --output emails.json
```

**Sample JSON structure:**
```json
{
  "timestamp": "2026-02-10T10:30:00.000000",
  "count": 2,
  "emails": [
    {
      "id": "AAMkAGE1NGQwYTk2LTkwYjEtNDB",
      "from": {
        "email": "vendor@example.com",
        "name": "Acme Suppliers"
      },
      "subject": "Invoice INV-2026-001 - Payment Due",
      "received_datetime": "2026-02-10T09:15:00Z",
      "body_preview": "Dear Planted, Please find invoice INV-2026-001 attached...",
      "body": "Dear Planted Delivery Team,\n\nI hope this email finds you well...",
      "has_attachments": true,
      "is_read": false,
      "importance": "high",
      "attachments": [
        {
          "id": "AAMkAGE1NGQwYTk2LTkwYjEtNDB",
          "name": "INV-2026-001.pdf",
          "content_type": "application/pdf",
          "size": 234567,
          "extracted_text": {
            "filename": "INV-2026-001.pdf",
            "content_type": "application/pdf",
            "text": "INVOICE\nInvoice Number: INV-2026-001\nDate: 2026-02-01\nAmount: CHF 5,000...",
            "extraction_method": "pdf_pdfplumber",
            "success": true
          }
        }
      ]
    }
  ]
}
```

**Using JSON output in Python:**
```python
import json

with open('emails.json') as f:
    data = json.load(f)
    
for email in data['emails']:
    print(f"{email['from']['name']}: {email['subject']}")
    for att in email['attachments']:
        print(f"  - {att['name']} ({att['size']} bytes)")
```

### Example 5: Detailed Text Report
```bash
python main.py read --days 3 --format detailed --output report.txt
```

**Sample structure:**
```
####################################################################################################
EMAIL #1
####################################################################################################

Subject:      Invoice INV-2026-001 - Payment Due
From:         Acme Suppliers <vendor@example.com>
Date:         2026-02-10T09:15:00Z
Priority:     HIGH
Read:         No
Has Attachments: Yes

----------------------------------------------------------------------------------------------------
FULL BODY:
----------------------------------------------------------------------------------------------------
Dear Planted Delivery Team,

I hope this email finds you well. As per our agreement, please find attached the invoice for February.

[Full email body text here...]

----------------------------------------------------------------------------------------------------
ATTACHMENTS (1):
----------------------------------------------------------------------------------------------------

[ATTACHMENT] INV-2026-001.pdf
  Type: application/pdf
  Size: 234567 bytes
  Extraction Method: pdf_pdfplumber
  Extracted Text:
  ----------------------------------------------------------------------------------------------------
  INVOICE
  
  Invoice Number: INV-2026-001
  Date: 2026-02-01
  Due Date: 2026-02-15
  
  Bill To: Planted AG
  ...
  ----------------------------------------------------------------------------------------------------
```

## Filtering & Searching

### Example 6: Search for Payment Reminders
```bash
python main.py read --search "subject:reminder OR subject:overdue" --format table
```

Shows only emails with "reminder" or "overdue" in subject line.

### Example 7: Find Unread Emails
```bash
python main.py read --search "isRead:false" --format json --output unread.json
```

Gets all unread emails and exports as JSON.

### Example 8: Filter by Sender
```bash
python main.py read --days 30 --search "from:bank@ubs.com"
```

Retrieves all emails from UBS bank (last 30 days).

### Example 9: Complex Search
```bash
python main.py read --search "(from:vendor OR from:supplier) AND subject:invoice" --days 14
```

Emails from vendors/suppliers mentioning invoice in last 14 days.

## Attachment Handling

### Example 10: Extract and Analyze Attachments
```bash
python main.py read --days 7 --format detailed --output detailed_with_attachments.txt
```

Full extraction of all attachments with text content.

### Example 11: Skip Attachments for Speed
```bash
python main.py read --days 30 --no-attachments --format json --output fast_export.json
```

Fast export without attachment processing. Use when you only need email metadata.

### Example 12: Specific Attachment Type
Currently, the tool extracts:
- **PDF**: Full text extraction
- **Excel**: Cell by cell content
- **CSV**: Direct file content
- **Images**: Placeholder (OCR planned for Phase 2)

To find emails with specific attachment types:
```bash
# Script using JSON output
python main.py read --days 7 --format json --output emails.json

# Python:
import json

with open('emails.json') as f:
    data = json.load(f)

for email in data['emails']:
    for att in email['attachments']:
        if att['content_type'] == 'application/pdf':
            print(f"PDF found in: {email['subject']}")
            if att.get('extracted_text') and att['extracted_text']['success']:
                print(f"  Content: {att['extracted_text']['text'][:200]}...")
```

## Advanced Scenarios

### Example 13: Analysis Script
```bash
# Export last 30 days
python main.py read --days 30 --format json --output last_month.json

# Then run analysis script (save as analyze_emails.py):
```

```python
#!/usr/bin/env python3
import json
from collections import Counter

with open('last_month.json') as f:
    data = json.load(f)

print(f"Total emails: {data['count']}")

senders = Counter()
importances = Counter()
attachment_count = 0

for email in data['emails']:
    senders[email['from']['email']] += 1
    importances[email['importance']] += 1
    attachment_count += len(email['attachments'])

print("\nTop senders:")
for sender, count in senders.most_common(5):
    print(f"  {sender}: {count} emails")

print("\nImportance distribution:")
for importance, count in importances.items():
    print(f"  {importance}: {count}")

print(f"\nTotal attachments: {attachment_count}")
```

**Run it:**
```bash
python analyze_emails.py
```

**Expected output:**
```
Total emails: 45

Top senders:
  vendor1@abc.com: 12 emails
  vendor2@xyz.com: 8 emails
  bank@ubs.com: 5 emails
  customer@example.com: 3 emails

Importance distribution:
  high: 8
  normal: 35
  low: 2

Total attachments: 23
```

### Example 14: Daily Report Generation
```bash
# Create daily_report.bat (Windows)
@echo off
set REPORTDIR=reports\%date:~-4,4%-%date:~-10,2%-%date:~-7,2%
mkdir %REPORTDIR% 2>nul

python main.py read --days 1 --format detailed --output %REPORTDIR%\daily_email_report.txt
python main.py read --days 1 --format json --output %REPORTDIR%\daily_email_data.json

echo Daily report generated in %REPORTDIR%
```

**Run daily via Task Scheduler:**
1. Open Task Scheduler
2. Create Basic Task
3. Set trigger: Daily at 8:00 AM
4. Set action: Run `C:\path\to\daily_report.bat`

### Example 15: Extract Specific Data
```python
import json
import re

with open('emails.json') as f:
    data = json.load(f)

# Extract invoice numbers
invoice_pattern = r'INV-\d+-\d+'

for email in data['emails']:
    matches = re.findall(invoice_pattern, email['subject'])
    if matches:
        print(f"{email['from']['email']}: {matches}")
    
    # Also check extracted text from attachments
    for att in email['attachments']:
        if att.get('extracted_text') and att['extracted_text']['success']:
            att_matches = re.findall(invoice_pattern, att['extracted_text']['text'])
            if att_matches:
                print(f"  Found in {att['name']}: {att_matches}")
```

## Automation

### Example 16: Windows Task Scheduler Setup

**Create a batch script** (`run_reader.bat`):
```batch
@echo off
cd C:\Users\MatthiasDiBianco\accounting-mailbox-reader
.\.venv\Scripts\python.exe main.py read --format json --output emails_%date:~-4,4%%date:~-10,2%%date:~-7,2%.json
```

**Schedule it:**
1. Open Task Scheduler
2. Create Basic Task
3. Name: "Read Accounting Emails"
4. Trigger: Weekly, Monday-Friday at 9:00 AM
5. Action: Start program `C:\path\to\run_reader.bat`
6. Add condition: Only if system is idle

### Example 17: Linux Cron Job

Add to crontab:
```bash
# Run every Monday-Friday at 9:00 AM
0 9 * * 1-5 cd /path/to/accounting-mailbox-reader && /path/to/venv/bin/python main.py read --format json --output emails_$(date +\%Y\%m\%d).json
```

### Example 18: Email Results

After setting up automation, process results:

```python
# process_results.py
import json
import os
from datetime import datetime

results_dir = "emails_*.json"

for filename in os.listdir("."):
    if filename.startswith("emails_") and filename.endswith(".json"):
        with open(filename) as f:
            data = json.load(f)
        
        # Process emails
        print(f"\nProcessing {filename} ({data['count']} emails)")
        
        for email in data['emails']:
            # Your processing logic here
            if email['importance'] == 'high':
                print(f"  [HIGH] {email['subject']} from {email['from']['email']}")
```

## Tips & Best Practices

### Performance
1. **Use `--no-attachments` for fast exports:** 10x faster when you don't need attachment content
2. **Filter by date range:** `--days 7` is faster than `--days 365`
3. **Use JSON for large datasets:** More efficient than table formatting

### Data Quality
1. **Always validate JSON:** Use `python -m json.tool emails.json` to check validity
2. **Test search queries:** Start simple, add complexity gradually
3. **Check extraction:** Use `--format detailed` to verify attachment content

### Scheduling
1. **Set up hourly/daily exports:** Keep historical data in version control
2. **Use consistent naming:** `emails_YYYYMMDD.json` for easy sorting
3. **Monitor logs:** Check `logs/accounting_triage.log` for errors

### Security
1. **Don't commit `.env`:** Already in `.gitignore`
2. **Limit file permissions:** Restrict access to exported JSON files
3. **Use strong client secret:** 24+ characters, rotated regularly
