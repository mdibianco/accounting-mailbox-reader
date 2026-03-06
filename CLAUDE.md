# Accounting Mailbox Reader — Claude Code Instructions

## Documentation Sync (Confluence)

**Every change to the tool's behavior MUST be reflected in Confluence documentation.**

The documentation lives at:
- **Confluence space:** Finance Department (`FC`)
- **Top page:** [Finance Helpdesk & Accounting Mailbox Agent](https://eatplanted.atlassian.net/wiki/spaces/FC/pages/2363719692)
- **Structure guide:** `docs/CONFLUENCE_STRUCTURE.md`

### When to update Confluence

| Change Type | Which Page(s) to Update |
|-------------|------------------------|
| New or modified email category | Top page (category table) + create/update the category's child page |
| Keyword rules changed (`keyword_rules.yaml`) | ⚙️ Technical Setup & Configuration |
| AI prompt changed (`classification_prompt.txt`, `pass2_ven_rem_prompt.txt`) | ⚙️ Technical Setup & Configuration |
| New entity added | Top page (entities table) + 🧾 VEN-INV (routing table) |
| Processing pipeline changed | 📥 How Emails Get Sorted (Triage) + ⚙️ Technical Setup |
| Outlook folder structure changed | 🔗 Email Chains & Conversations + ⚙️ Technical Setup |
| Dashboard / workflow changed | 📊 Dashboard & Daily Workflow |
| Conversation matching logic changed | 🔗 Email Chains & Conversations |
| Scheduling or automation changed | ⚙️ Technical Setup & Configuration |
| New rigid rule / pre-filter added | 📥 How Emails Get Sorted + 🔇 NO_ACTION_NEEDED & OTHER |

### How to update

Use the Atlassian MCP tools:
- `mcp__claude_ai_Atlassian__getConfluencePage` to read existing content
- `mcp__claude_ai_Atlassian__updateConfluencePage` to update (cloudId: `eatplanted.atlassian.net`, contentFormat: `markdown`)
- `mcp__claude_ai_Atlassian__createConfluencePage` to create new pages (spaceId: `99942446`, parentId: `2363719692`)

### Page IDs reference

| Page | ID |
|------|----|
| Top page (overview) | 2363719692 |
| How Emails Get Sorted (Triage) | 2551021593 |
| VEN-REM — Vendor Payment Reminders | 2549284918 |
| VEN-INV — Vendor Invoices | 2551218197 |
| VEN-FOLLOWUP — Vendor Queries | 2551087116 |
| CUST-REM-FOLLOWUP & CUST-REMIT | 2550464556 |
| NO_ACTION_NEEDED & OTHER | 2549809205 |
| Email Chains & Conversations | 2550595617 |
| Dashboard & Daily Workflow | 2550824980 |
| Technical Setup & Configuration | 2549317666 |

## Project Conventions

- **Mailbox:** accounting@eatplanted.com
- **Entity mapping:** CH1 → invoices@eatplanted.com, others → `{code.lower()}-invoices@eatplanted.com`
- **JSON backward compat:** Power Query uses explicit field enumeration — adding fields is safe, removing is not
- **AI models:** Google Gemini free tier with cascade fallback
- **Draft replies:** Saved (not auto-sent) per user preference
