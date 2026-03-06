# Email Categories

Every email that arrives in the accounting mailbox is sorted into one of these categories:

## VEN-REM - Vendor Payment Reminder

A vendor or supplier is reminding us to pay an outstanding bill.

**Examples:** "Zahlungserinnerung", "Mahnung", "Payment reminder", "Sollecito di pagamento"

**What happens:** These get the full treatment - the tool reads the email in detail, extracts invoice numbers and amounts, checks urgency, and translates if needed. After processing, the email moves to ARCHIVE / PROCESSED BY AGENT.

**What you do:** Review in the dashboard. Check if it's already being handled, then pay or reply.

## VEN-INV - Vendor Invoice

A vendor has sent us a new invoice or bill.

**Examples:** "Rechnung", "Invoice attached", "Faktura", "Facture"

**What happens:** Sorted and translated, but no detailed extraction yet (planned for the future).

**What you do:** Process as usual - book the invoice in Business Central.

## CUST-REM-FOLLOWUP - Customer Follow-up

A customer is replying to a payment reminder we sent them.

**Examples:** "AW: Planted Foods AG - Issued Reminder CREP01002118", "Re: Outstanding invoice"

**What happens:** Sorted and translated. Often these are auto-replies, out-of-office messages, or payment promises.

**What you do:** Check what the customer said. If they dispute an invoice, escalate. If they promise to pay, note the date.

## CUST-REMIT - Customer Remittance Advice

A customer confirms they have paid us (or are about to).

**Examples:** "Remittance advice", "Zahlungsavis", "Payment notification"

**What happens:** Sorted and translated.

**What you do:** Match the payment against open receivables in Business Central.

## OTHER - Everything else

Emails that don't fit the above categories: internal forwards, newsletters, IT notifications, bounced emails, out-of-office replies, etc.

**What happens:** Sorted with low priority. No further processing.

**What you do:** Nothing in most cases. Check occasionally for emails that were wrongly categorised.

## Priority levels

Each email also gets a priority tag in Outlook:

| Priority | Meaning |
|----------|---------|
| **PRIO_HIGHEST** | Legal threats, final warnings, very urgent payment demands |
| **PRIO_HIGH** | Standard payment reminders, overdue invoices |
| **PRIO_MEDIUM** | Regular invoices, confirmations, everything else |

## Email chains

When multiple emails belong to the same conversation (e.g. a vendor sends 3 reminders over 2 weeks), the tool links them together. In the dashboard, you see the full chain with the most recent email on top.

This also means you won't see the same vendor reminder 5 times in your to-do list. Only the latest one in each chain shows as active.
