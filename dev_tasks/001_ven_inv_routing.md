# VEN-INV: Repeat Offender Tracking & Nudge Emails

**Status:** BACKLOG
**Category:** VEN-INV

## Already Implemented

- Forward to correct entity invoices address (enabled, working)
- Reply to sender with correct invoices address
- CC @eatplanted.com employees found in the email chain
- Entity detection (fuzzy name + VAT pattern matching)
- Mailbox lookup (3-day sender+subject dedup)
- Draft reply with correct address

## Remaining Work

1. **Repeat offender tracking**: Flag senders who repeatedly send to accounting@ instead of invoices@
   - Needs a persistent store (YAML or JSON counter per sender)
   - Frequency analysis and escalation rules
2. **Planted responsible extraction**: From context, attachments, or email chain, identify the internal contact responsible for that vendor
3. **Draft nudge email** to the planted responsible urging them to update the vendor's invoicing address
