# Processing Statistics - Power Query Guide

## Overview
The accounting mailbox reader now tracks all process and cleanup runs. These statistics are saved as JSON files in the emails folder for Power BI dashboard integration.

## JSON Files Location
- **Process Runs**: `C:\Users\MatthiasDiBianco\emails\emails\process_runs.json`
- **Cleanup Runs**: `C:\Users\MatthiasDiBianco\emails\emails\cleanup_runs.json`

## Data Structure

### Process Runs (process_runs.json)
Each record contains:
```json
{
  "timestamp": "2026-03-04T10:30:15.123456Z",
  "type": "scheduled_run",
  "emails_processed": 5,
  "categories_by_keywords": 3,
  "categories_by_llm": 2,
  "ven_rem_analyzed": 2,
  "ven_followup_analyzed": 0,
  "ven_inv_processed": 1,
  "emails_archived": 3,
  "human_completed_moved": 2,
  "jsons_saved": 5,
  "llm_calls_by_model": {
    "gemini-2.5-flash": 2
  }
}
```

### Cleanup Runs (cleanup_runs.json)
Each record contains:
```json
{
  "timestamp": "2026-03-04T17:00:45.123456Z",
  "type": "cleanup_run",
  "emails_in_folder": 120,
  "emails_classified": 85,
  "emails_with_pass2": 42,
  "emails_moved_to_archive": 38,
  "api_calls_used": 42
}
```

## Power Query Integration

### Setup in Power BI Desktop

1. **For Process Runs**:
   - New > Query > New Blank Query
   - Copy contents of `PQ_ProcessRuns.m`
   - Paste into Advanced Editor
   - Load the query as "ProcessRuns" table

2. **For Cleanup Runs**:
   - New > Query > New Blank Query
   - Copy contents of `PQ_CleanupRuns.m`
   - Paste into Advanced Editor
   - Load the query as "CleanupRuns" table

### Using in Dashboard

Once loaded, you can create visualizations:

**ProcessRuns Table**:
- Daily email processing summary
- Keyword vs LLM categorization trends
- Category-specific processing counts (VEN-REM, VEN-FOLLOWUP, VEN-INV)
- LLM API call usage by model
- Archive efficiency metrics

**CleanupRuns Table**:
- Daily cleanup summary
- Pass 2 analysis completion rate
- Archive efficiency
- API budget usage per cleanup run

### Sample Calculations

```m
// Daily totals (in new column)
Daily Emails = SUMX(FILTER(ProcessRuns, ProcessRuns[run_date] = TODAY()), ProcessRuns[emails_processed])

// Keyword vs LLM split
Keywords Pct = DIVIDE(
    SUM(ProcessRuns[categories_by_keywords]),
    SUM(ProcessRuns[emails_processed])
)

// API cost tracking
Total API Calls = SUMX(CleanupRuns, CleanupRuns[api_calls_used])

// Archive efficiency
Archive Rate = DIVIDE(
    SUM(ProcessRuns[emails_archived]),
    SUM(ProcessRuns[emails_processed])
)
```

## Data Retention
- **Process Runs**: Last 365 days (auto-trimmed)
- **Cleanup Runs**: Last 90 days (auto-trimmed)

## Refresh Schedule
- Process runs logged: Hourly (08:00-17:00, plus 17:00 cleanup)
- Data available for BI refresh immediately after each run
- Recommend Power BI refresh schedule: Every 1 hour or on demand

## Troubleshooting

**Power Query can't find JSON file**:
- Verify file path matches `C:\Users\MatthiasDiBianco\emails\emails\`
- Check that emails folder has write permissions
- Ensure at least one process run has completed

**JSON structure errors**:
- Check that daily_stats.py version is current
- Look for errors in main.py log file
- Verify JSON syntax with `python -m json.tool process_runs.json`

**Data not updating**:
- Verify process runs completed successfully
- Check that `--upload-sharepoint` flag was used in batch command
- Refresh Power BI query in Power Query Editor
