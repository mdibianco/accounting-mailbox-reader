# Plan Queue

## 1. VEN-INV: Full Invoice Routing Pipeline

**Goal**: End-to-end handling of vendor invoices sent to accounting@ instead of the correct invoices mailbox.

### Steps
1. **Forward** to correct entity invoices address (already partially implemented, currently disabled)
2. **Reply to sender** with the correct invoices address for future submissions
3. **CC planted employees** found in the email chain (CC list or names on the invoice)
4. **Repeat offender tracking**: Flag senders who repeatedly send to accounting@ instead of invoices@
5. **Planted responsible extraction**: From context, attachments, or email chain, identify the internal contact responsible for that vendor
6. **Draft nudge email** to the planted responsible urging them to update the vendor's invoicing address

### Notes
- Forwarding was disabled (line 96 in `src/pass2_inv_processor.py`) — needs fix before re-enabling
- Entity detection is currently fuzzy text search; may need LLM assist for ambiguous cases
- Repeat offender logic needs a persistent store (YAML or JSON counter per sender)


## 2. CUST-REM-FOLLOWUP: Customer Matching & Reminder Extraction

**Goal**: Match incoming customer reminder/follow-up emails to customer master data so they appear correctly in the dashboard. Extract which specific reminder is being challenged.

### Matching Routine (cascading, LLM-light)
1. **Exact match** on email domain vs customer domain
2. **Trim/fuzzy match** on sender name vs customer name
3. **LLM fallback** (only if 1+2 fail) to extract invoice numbers, customer references, etc.

### Extraction
- Identify which specific reminder level or invoice is being disputed/queried
- Structure for dashboard merge with customer record

### Notes
- Should be very LLM-light — prioritize deterministic matching
- Customer master data source TBD (BC export? Graph API?)


## 3. CUST-PAYM: Customer Payment Advices (NEW CATEGORY)

**Goal**: New category for customer payment notifications. Many already caught by office-specific keywords; formalize as a proper category with its own Pass 2.

### Classification
1. **Email address rules**: Many sender addresses automatically map to CUST-PAYM
2. **Keyword rules**: Payment-related keywords in subject/body
3. **LLM fallback**: For ambiguous cases

### Pass 2: Structured Payment Extraction
- Per-customer parser that structures payment advice attachments into a BC-uploadable format
- **Start with Edeka** as first customer implementation
- Extract: invoice numbers, amounts, payment dates, deductions/offsets

### Notes
- Currently these emails are caught by generic office keyword rules
- Each customer may have a different attachment format (PDF, CSV, EDI)
- BC upload format specification needed before building parsers
