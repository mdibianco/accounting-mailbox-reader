# 📚 Confluence Documentation Structure

> **Space:** Finance Department (`FC`)
> **Top Page:** [Finance Helpdesk & Accounting Mailbox Agent](https://eatplanted.atlassian.net/wiki/spaces/FC/pages/2363719692)

## Page Tree

```
📄 Finance Helpdesk & Accounting Mailbox Agent          ← Overview, links to all sub-pages
│
├── 📄 How Emails Get Sorted (Triage)                   ← Pass 0 keywords → Pass 1 AI → confidence levels
│
├── 📄 VEN-REM — Vendor Payment Reminders               ← Agent extracts details, human reviews & pays
│
├── 📄 VEN-INV — Vendor Invoices                        ← Agent detects entity & forwards, human books invoice
│
├── 📄 VEN-FOLLOWUP — Vendor Queries                    ← Human must respond (netting, remittance, etc.)
│
├── 📄 CUST-REM-FOLLOWUP & CUST-REMIT — Customer Emails← Human matches payments / handles disputes
│
├── 📄 NO_ACTION_NEEDED & OTHER                         ← Auto-filtered noise + uncategorized emails
│
├── 📄 Email Chains & Conversations                     ← Threading, supersession, CHAIN category
│
├── 📄 Dashboard & Daily Workflow                        ← Power BI dashboard, daily routine, marking done
│
└── 📄 Technical Setup & Configuration                   ← Architecture, keywords, prompts, scheduling
```

## Audience

| Page | Primary Audience | Purpose |
|------|-----------------|---------|
| Overview | Everyone | What is this tool, quick start |
| Triage | Team | Understand how sorting works |
| VEN-REM | AP Team | Process vendor payment reminders |
| VEN-INV | AP Team | Handle incoming invoices |
| VEN-FOLLOWUP | AP Team | Respond to vendor questions |
| Customer Emails | AR Team | Handle customer replies & remittances |
| NO_ACTION_NEEDED & OTHER | Team | Know what's filtered, spot misclassifications |
| Email Chains | Team | Understand conversation grouping |
| Dashboard & Workflow | Team | Daily routine and Power BI |
| Technical Setup | Admin / IT | Architecture, config, troubleshooting |

## Maintenance

- **Keyword rules** change → update Technical Setup page
- **New category** added → create a new child page + update Overview
- **Workflow changes** → update Dashboard & Daily Workflow page
- **Architecture changes** → update Technical Setup page
