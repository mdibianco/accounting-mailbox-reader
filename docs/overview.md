# Accounting Mailbox Reader

## What is this?

The Accounting Mailbox Reader checks the **accounting@eatplanted.com** inbox every hour, reads each email, sorts it into a category, and for payment reminders extracts the key details (who is asking, which invoices, how urgent).

The results feed into a Power BI dashboard where you can see everything at a glance.

## What changes for you?

### What you no longer need to do

- **Read every email to figure out what it is.** The tool sorts vendor reminders, invoices, customer confirmations, etc. automatically.
- **Manually flag or tag emails.** Each email gets an Outlook category tag (e.g. "VEN-REM", "PRIO_HIGH") automatically.
- **Translate German/French/Italian emails.** Every non-English email gets an English summary and full translation.
- **Track which reminder is the 2nd or 3rd follow-up.** The tool links email chains together so you see the full history.
- **Open Business Central to check if an invoice is paid.** *(coming soon)* The tool will look this up for you.

### What you still need to do

- **Review the dashboard regularly** to see new vendor reminders and their urgency.
- **Take action** on open items: pay invoices, reply to vendors, escalate disputes.
- **Mark emails as done in Outlook when you've handled them.** This is important: complete the flag on the email (right-click > Follow Up > Mark Complete) so the tool knows it's been taken care of. On the next run, the tool picks up your completed flag and moves the email to **ARCHIVE / PROCESSED BY HUMANS**. If you don't flag it as done, the email stays in the inbox and keeps showing up as open.
- **Report any misclassifications** so the sorting rules can be improved.

### What the tool does NOT do

- It does **not** reply to any emails.
- It does **not** make payments or change anything in Business Central.
- It does **not** delete any emails.
- It is read-only from a business perspective.

## Limits

- The tool uses a free AI service (Google Gemini) with a daily limit of ~60 calls. This is enough for normal daily email volume (~15-20 new emails). If there's a large backlog, it processes the most important ones first and catches up over several days.
- Attachment text extraction works for PDF, Excel, and images but is not perfect. Large or scanned PDFs may be partially extracted.
- The sorting is correct ~95% of the time. Unusual emails or mixed-content emails may be miscategorised.

## How it runs

| Time | What happens |
|------|-------------|
| 08:00 - 16:00 | Runs every hour. Picks up new emails, sorts them, analyses vendor reminders, saves results. |
| 17:00 | Last run of the day. After normal processing, uses any remaining AI budget to work through the Reminders backlog folder. |
| 18:00 - 07:00 | Paused. No processing overnight. |

The log file is at: `C:\Users\MatthiasDiBianco\.accounting_mailbox_reader\process.log`
