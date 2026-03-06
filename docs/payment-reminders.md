# Payment Reminders (VEN-REM)

## Why these get special treatment

Vendor payment reminders used to be the most time-consuming emails to process. For each one, you had to:

1. Read the email (often in German, French, or Italian)
2. Figure out which invoices they're asking about
3. Note the invoice numbers, amounts, and due dates
4. Open Business Central to check if we've already paid
5. Decide how urgent it is

The tool now does steps 1-3 automatically. Step 4 (BC lookup) is planned but currently disabled.

## What the tool extracts

For every payment reminder, you'll see in the dashboard:

| Field | Example |
|-------|---------|
| **Vendor** | Nordfrost, DPD, Crayon |
| **Our entity** | CH1 (Planted Foods AG), DE1 (Planted Foods GmbH), etc. |
| **Urgency** | Low (first reminder), Medium (Mahnung), High (final warning) |
| **Invoice numbers** | RE-2024-001, 4212511001974 |
| **Amounts** | EUR 5,000.00, CHF 1,234.56 |
| **Due dates** | 15.01.2026 |
| **English summary** | 2-3 sentences explaining who wants what |
| **Link to Outlook** | Click to open the original email |

## Example

An email arrives from DPD:

> **Subject:** D.0193 Mahnung 002300184445
>
> Sehr geehrte Damen und Herren, trotz unserer Erinnerung ist die Rechnung Nr. 002300184445 uber EUR 1.234,56 vom 15.01.2026 noch nicht beglichen...

The tool produces:
- **Vendor:** DPD
- **Entity:** DE1 (Planted Foods GmbH)
- **Urgency:** Medium (Mahnung = second-level reminder)
- **Invoice:** 002300184445, EUR 1,234.56, due 15.01.2026
- **Summary:** "DPD is sending a payment reminder for invoice 002300184445 (EUR 1,234.56, due Jan 15 2026). This is a second-level reminder."

The email is moved to **ARCHIVE / PROCESSED BY AGENT** and appears in the dashboard.

## What you do

1. Open the Power BI dashboard
2. Look at the open vendor reminders, sorted by urgency
3. For each one: check if the invoice is already scheduled for payment
4. If not paid: initiate payment or contact the vendor
5. If already paid: the vendor's records may be outdated, you can let them know
6. In Outlook: complete the flag on the email when done. The tool will move it to PROCESSED BY HUMANS.

## Conversation chains

If a vendor sends multiple reminders (e.g. first reminder, second reminder, Mahnung, final warning), these are grouped together. You only see the most recent one as active. The older ones are linked as history so you can see the full escalation.
