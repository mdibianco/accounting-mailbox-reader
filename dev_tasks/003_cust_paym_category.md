# CUST-PAYM: Customer Payment Advices (NEW CATEGORY)

**Status:** BACKLOG
**Category:** CUST-PAYM

## Goal

New category for customer payment notifications. Many already caught by office-specific keywords; formalize as a proper category with its own Pass 2.

## Classification

1. **Email address rules**: Many sender addresses automatically map to CUST-PAYM
2. **Keyword rules**: Payment-related keywords in subject/body
3. **LLM fallback**: For ambiguous cases

## Pass 2: Structured Payment Extraction

- Per-customer parser that structures payment advice attachments into a BC-uploadable format
- **Start with Edeka** as first customer implementation
- Extract: invoice numbers, amounts, payment dates, deductions/offsets

## Notes

- Currently these emails are caught by generic office keyword rules
- Each customer may have a different attachment format (PDF, CSV, EDI)
- BC upload format specification needed before building parsers
