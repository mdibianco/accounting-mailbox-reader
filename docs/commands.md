# Commands Reference

All commands are run from the project folder:

```
cd c:\Users\MatthiasDiBianco\accounting-mailbox-reader
.venv\Scripts\activate
```

---

## Daily operations

### process

**The main command. Runs automatically every hour - you rarely need to run this manually.**

```
python main.py process --upload-sharepoint
```

Fetches only new emails (since last run), classifies them, analyses vendor reminders, translates, archives, and saves JSONs.

Options:
- `--upload-sharepoint` - save email JSONs to the local folder (required for Power BI)
- `--dry-run` - preview what would happen without changing anything in Outlook
- `--mailbox EMAIL` - use a different mailbox (default: accounting@eatplanted.com)

### cleanup

**Uses remaining daily AI budget to process older emails from a specific folder.**

```
python main.py cleanup --upload-sharepoint
```

Runs automatically at 17:00 after the last `process` run. Targets the Reminders folder by default.

Options:
- `--folder NAME` - which Outlook folder to process (default: "Reminders")
- `--budget N` - max AI calls to use (default: whatever's left today)
- `--days N` - only process emails from last N days (default: 60)
- `--dry-run` - preview only

---

## Manual / one-off commands

### read

**Read and process emails manually with full control over options.**

```
python main.py read --days 7 --max 50 --classify --deep --upload-sharepoint --write-back
```

Options:
- `--days N` - how many days back (default: 7)
- `--max N` - max emails to read (default: 50)
- `--classify` - run AI classification on uncertain emails
- `--deep` - run Pass 2 analysis on vendor reminders (implies --classify)
- `--force-llm` - force AI classification on ALL emails (ignores keyword results)
- `--write-back` - write category tags back to Outlook
- `--upload-sharepoint` - save JSONs
- `--format json|table|detailed` - output format (default: table)
- `--output FILE` - save output to file
- `--no-attachments` - skip attachment extraction
- `--no-body` - skip full body content

### preview

**Quick look at the 5 most recent emails without any processing.**

```
python main.py preview
```

---

## Setup and maintenance

### config-show

**Show current configuration and whether everything is connected.**

```
python main.py config-show
```

### init

**Create the .env file from the template (first-time setup only).**

```
python main.py init
```

### sync-categories

**Re-sync email categories from the Confluence page.**

```
python main.py sync-categories
```

### build-conversation-index

**Rebuild the conversation index from all existing email JSONs.**

```
python main.py build-conversation-index
```

Run this if the conversation index gets corrupted or if you want to rebuild it from scratch. Normally the index maintains itself automatically.

### test-classify

**Test AI classification on a single email (debugging).**

```
python main.py test-classify
```
