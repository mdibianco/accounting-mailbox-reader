# Architecture

## Processing pipeline

Each email goes through these steps in order:

```
Email arrives
  |
  v
Pass 0: Keyword matching (instant, no AI cost)
  - Scans subject + body for known terms in 6 languages
  - Assigns category + confidence (HIGH / MEDIUM / LOW)
  |
  v
Conversation matching
  - Strips RE:/FW:/AW: prefixes from subject
  - Searches existing emails (last 30 days) for matching subjects
  - Links email chains together
  |
  v
Pass 1: AI classification (only LOW-confidence emails, ~20-30%)
  - Sends email to Google Gemini for a second opinion
  - Costs 1 API call per email
  |
  v
Pass 2: Deep analysis (VEN-REM only)
  - Extracts vendor name, entity, urgency, invoice details
  - Costs 1 API call per email
  - BC lookup currently disabled
  |
  v
Translation (non-English, non-Pass-2 emails)
  - English summary + full body translation
  - Costs 1 API call per email
  |
  v
Save + Archive
  - JSON file saved to local folder (picked up by Power BI)
  - Outlook: category tags, read flag, extended properties
  - VEN-REM with Pass 2: moved to ARCHIVE / PROCESSED BY AGENT
```

## AI model cascade

The tool uses Google Gemini's free tier. When one model hits its daily limit, it falls back to the next:

1. gemini-2.5-flash (fastest)
2. gemini-3-flash-preview (fallback)
3. gemini-2.5-flash-lite (last resort)

Daily budget: ~60 API calls across all models.

## File structure

```
emails/emails/                      Local folder with one JSON per email
  ├── 2026-03-01_15-31-31Z_d7b6.json
  ├── 2026-03-02_07-14-45Z_75de.json
  ├── ...
  └── conversation_index.json       Lightweight index for conversation matching
```

Each email JSON contains: sender, subject, body, attachments (with extracted text), classification, Pass 2 results, English translation, conversation links.

JSON size is capped at 200 KB per file. Email bodies are truncated at 10,000 chars, attachment text at 15,000 chars per attachment.

## Scheduled automation

Windows Task Scheduler runs `run_process.bat` hourly 08:00-17:00:
- Normal `process` command (new emails only, using watermark)
- At 17:00: additional `cleanup` command for the Reminders folder backlog

Log file: `C:\Users\MatthiasDiBianco\.accounting_mailbox_reader\process.log` (auto-cleared after 30 days)

API call tracking: `C:\Users\MatthiasDiBianco\.accounting_mailbox_reader\api_calls.json`

## Key configuration

All settings in `.env` file:
- `ACCOUNTING_MAILBOX` - target mailbox address
- `LOCAL_FOLDER_PATH` - where email JSONs are saved
- `AZURE_TENANT_ID`, `AZURE_CLIENT_ID` - for Outlook access via Microsoft Graph API
- `GEMINI_API_KEY` - optional, falls back to CLI OAuth if not set

## Outlook folder structure

```
accounting@eatplanted.com
  ├── Inbox                         New emails arrive here
  ├── Reminders                     Backlog folder (processed at 17:00)
  └── ARCHIVE
      ├── PROCESSED BY AGENT        Emails fully handled by the tool
      └── PROCESSED BY HUMANS       Emails marked done by the team
```
