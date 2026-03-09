# CUST-REM-FOLLOWUP: Reminder Number & Domain Extraction

**Status:** DONE
**Category:** CUST-REM-FOLLOWUP

## Goal

Extract structured data from CUST-REM-FOLLOWUP emails in Pass 2:
1. Reminder number from the subject (pattern: `CREP...`)
2. Sender's email domain as customer identifier

Save results in JSON for dashboard use.

## Steps

1. Create `src/pass2_cust_rem_processor.py`
   - Extract reminder number: regex for `CREP\w+` pattern in subject
   - Extract domain from sender email address (fallback identifier)
   - Return structured results for JSON output
2. Add pass2_results fields to `Email.to_dict()` in `src/email_reader.py`
   - `cust_reminder_number` (string or null)
   - `cust_domain` (string)
3. Wire into `main.py` Pass 2 block for CUST-REM-FOLLOWUP emails
4. Add columns to `powerquery/PQ_Emails.m`

## Notes

- No LLM needed — pure regex + string extraction
- Category already defined and classified by LLM in Pass 1
- CREP pattern needs verification against real email subjects
