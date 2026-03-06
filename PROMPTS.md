# CLI Commands & Prompts

## Scheduled / Production

These are the commands that run automatically via Task Scheduler (or Docker cron).

```bash
# Full automated run: fetch new emails, classify, deep-analyze, archive, save JSONs
python main.py process --upload-sharepoint

# Cleanup run (17:00): process Reminders folder with remaining daily API budget
python main.py cleanup --upload-sharepoint

# Dry run (see what would happen without Outlook changes)
python main.py process --upload-sharepoint --dry-run
python main.py cleanup --dry-run
```

---

## Interactive / Manual

### Read & classify emails (ad-hoc)

```bash
# Read last 7 days, table output (default)
python main.py read

# Read with full AI classification + JSON export
python main.py read --classify --format json --output classified.json

# Deep analysis (implies --classify): keyword triage + LLM + Pass 2 on VEN-REM
python main.py read --deep --upload-sharepoint

# Force LLM on ALL emails (ignores keyword confidence)
python main.py read --force-llm --format json

# Write categories back to Outlook (tags + extended properties)
python main.py read --classify --write-back

# Custom date range and batch size
python main.py read --days 14 --max 100

# Search specific emails (OData syntax)
python main.py read --search "from:vendor@example.com"

# Skip attachments (faster, no PDF text)
python main.py read --no-attachments
```

### Cleanup with custom options

```bash
# Use exactly 19 API calls
python main.py cleanup --budget 19

# Process a different folder
python main.py cleanup --folder "OtherFolder"

# Process older emails (default: 60 days)
python main.py cleanup --days 90
```

### Quick preview

```bash
# Preview 5 most recent emails (no classification, no changes)
python main.py preview
```

### Test & debug

```bash
# Test classification on the most recent email
python main.py test-classify

# Show current configuration (credentials, paths, etc.)
python main.py config-show

# Sync categories from Confluence
python main.py sync-categories

# Build/rebuild conversation index from existing JSONs
python main.py build-conversation-index

# Initialize .env from template
python main.py init
```
