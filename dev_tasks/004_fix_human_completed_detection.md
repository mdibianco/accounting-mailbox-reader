# Fix Human-Completed Email Detection (Shared Mailbox Flag Bug)

**Status:** DONE

## Problem

Outlook flag status (`complete` / green tick) set by users on the shared mailbox (accounting@eatplanted.com) is NOT visible to Graph API. Flags on shared mailboxes are stored per-user in Outlook, not on the server-side message.

## Solution (Implemented)

Switched to Outlook categories. `get_inbox_messages_by_category(mailbox, "DONE")` in `src/graph_client.py` (lines 532-569). Step 10 in `main.py` uses dual detection: category "DONE" (primary) + flag "complete" (fallback), deduped by message ID.

## Remaining

- Confluence docs update (Dashboard & Daily Workflow page 2550824980, Technical Setup page 2549317666)
