"""Conversation matching: group related emails (RE/FW chains, repeated reminders).

Approach:
1. Primary: Match on Graph API conversationId (Outlook's native threading)
2. Fallback: Strip subject prefixes and match normalized subject
3. Last resort: Levenshtein on normalized subject (same sender domain only)
4. Only look back 30 days — older conversations are considered dead

The conversation index is a lightweight JSON file that stores entries by both
graph_conversation_id and stripped_subject, so we never need to re-read
hundreds of email JSONs on each run.
"""

import json
import logging
import re
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)

# How far back to look for conversation matches
CONVERSATION_WINDOW_DAYS = 30

# Levenshtein similarity threshold for subject matching (last resort)
SUBJECT_SIMILARITY_THRESHOLD = 0.85

# Multilingual reply/forward prefixes to strip (order matters: longer first)
_SUBJECT_PREFIXES = [
    # Auto-replies / bounces (strip entire prefix up to the real subject)
    r"Risposta automatica:\s*",
    r"Automatische Antwort:\s*",
    r"Automatic reply:\s*",
    r"Abwesend\s*/\s*\S+\s+",            # "Abwesend / Ferien Re:"
    r"Undeliverable:\s*",
    r"Non recapitabile:\s*",
    r"Non remis:\s*",
    # Forward prefixes
    r"Fwd:\s*",
    r"FW:\s*",
    r"WG:\s*",                             # German: Weitergeleitet
    r"I:\s*",                              # Italian: Inoltrato
    r"TR:\s*",                             # French: Transféré
    r"VS:\s*",                             # Finnish: Välitetty
    r"VB:\s*",                             # Swedish: Vidarebefordrat
    r"Doorgestuurd:\s*",                   # Dutch
    # Reply prefixes
    r"RE:\s*",
    r"Re:\s*",
    r"AW:\s*",                             # German: Antwort
    r"R:\s*",                              # Italian: Risposta
    r"SV:\s*",                             # Swedish/Norwegian/Danish: Svar
    r"Antw:\s*",                           # Dutch: Antwoord
    # Internal system prefixes
    r"HOMA6:\s*",                          # Planted's Issued Reminder system
]

# Compile into one pattern that strips all prefixes (repeatedly)
_PREFIX_PATTERN = re.compile(
    r"^(?:" + "|".join(_SUBJECT_PREFIXES) + r")+",
    re.IGNORECASE,
)


def normalize_subject(subject: str) -> str:
    """Strip reply/forward prefixes and normalize whitespace + case."""
    if not subject:
        return ""
    stripped = _PREFIX_PATTERN.sub("", subject)
    # Normalize whitespace and lowercase
    return re.sub(r"\s+", " ", stripped).strip().lower()


def levenshtein_similarity(s1: str, s2: str) -> float:
    """Levenshtein distance normalized to 0-1 similarity.
    Quick-rejects if lengths differ by >20%."""
    len1, len2 = len(s1), len(s2)
    max_len = max(len1, len2)
    if max_len == 0:
        return 1.0
    if abs(len1 - len2) / max_len > 0.2:
        return 0.0

    # Standard DP
    prev = list(range(len2 + 1))
    for i in range(len1):
        curr = [i + 1]
        for j in range(len2):
            cost = 0 if s1[i] == s2[j] else 1
            curr.append(min(curr[j] + 1, prev[j + 1] + 1, prev[j] + cost))
        prev = curr

    return 1.0 - prev[len2] / max_len


class ConversationIndex:
    """Maintains a lightweight index for conversation matching.

    Index file structure:
    {
        "last_updated": "2026-02-27T18:00:00Z",
        "by_graph_id": {
            "AAQk...==": {
                "conversation_id": "conv_abc123",
                "entries": [
                    {
                        "filename": "2026-01-07_07-59-30Z_abc123.json",
                        "date": "2026-01-07T07:59:30Z",
                        "sender_domain": "nordfrost.de",
                        "original_subject": "Zahlungserinnerung",
                        "processing_status": "ARCHIVE > PROCESSED BY AGENT"
                    }
                ]
            }
        },
        "by_subject": {
            "normalized subject text": {
                "conversation_id": "conv_abc123",
                "graph_id": "AAQk...==",
                "entries": [...]
            }
        }
    }
    """

    def __init__(self, email_folder: str):
        self.email_folder = Path(email_folder)
        self.index_path = self.email_folder / "conversation_index.json"
        self.by_graph_id: dict[str, dict] = {}
        self.by_subject: dict[str, dict] = {}
        self._load()

    def _load(self):
        """Load existing index from disk."""
        if self.index_path.exists():
            try:
                with open(self.index_path, "r", encoding="utf-8") as f:
                    data = json.load(f)

                # Support new format (by_graph_id + by_subject)
                if "by_graph_id" in data:
                    self.by_graph_id = data.get("by_graph_id", {})
                    self.by_subject = data.get("by_subject", {})
                else:
                    # Migrate from old format (flat "entries" dict keyed by subject)
                    self._migrate_old_format(data.get("entries", {}))

                total = sum(len(g["entries"]) for g in self.by_graph_id.values())
                total += sum(
                    len(s["entries"]) for s in self.by_subject.values()
                    if "graph_id" not in s  # Don't double-count
                )
                logger.info(
                    f"Loaded conversation index: {total} entries "
                    f"({len(self.by_graph_id)} graph IDs, {len(self.by_subject)} subjects)"
                )
            except (json.JSONDecodeError, KeyError) as e:
                logger.warning(f"Corrupt conversation index, rebuilding: {e}")
                self.by_graph_id = {}
                self.by_subject = {}
        else:
            logger.info("No conversation index found — will be created on save")

    def _migrate_old_format(self, old_entries: dict):
        """Migrate from old subject-only index to new dual-key format."""
        logger.info("Migrating conversation index from old format...")
        self.by_graph_id = {}
        self.by_subject = {}
        for norm_subj, entry_list in old_entries.items():
            if not entry_list:
                continue
            conv_id = entry_list[0].get("conversation_id", _generate_conversation_id(norm_subj))
            self.by_subject[norm_subj] = {
                "conversation_id": conv_id,
                "entries": entry_list,
            }

    def save(self):
        """Save index to disk."""
        data = {
            "last_updated": datetime.now(timezone.utc).isoformat(),
            "by_graph_id": self.by_graph_id,
            "by_subject": self.by_subject,
        }
        self.email_folder.mkdir(parents=True, exist_ok=True)
        with open(self.index_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        total_g = sum(len(g["entries"]) for g in self.by_graph_id.values())
        total_s = sum(len(s["entries"]) for s in self.by_subject.values() if "graph_id" not in s)
        logger.info(f"Saved conversation index: {total_g + total_s} entries")

    def prune(self, days: int = CONVERSATION_WINDOW_DAYS):
        """Remove entries older than N days."""
        cutoff = (datetime.now(timezone.utc) - timedelta(days=days)).isoformat()
        pruned = 0

        # Prune by_graph_id
        empty_keys = []
        for gid, group in self.by_graph_id.items():
            before = len(group["entries"])
            group["entries"] = [e for e in group["entries"] if e.get("date", "") >= cutoff]
            pruned += before - len(group["entries"])
            if not group["entries"]:
                empty_keys.append(gid)
        for key in empty_keys:
            del self.by_graph_id[key]

        # Prune by_subject
        empty_keys = []
        for subj, group in self.by_subject.items():
            before = len(group["entries"])
            group["entries"] = [e for e in group["entries"] if e.get("date", "") >= cutoff]
            pruned += before - len(group["entries"])
            if not group["entries"]:
                empty_keys.append(subj)
        for key in empty_keys:
            del self.by_subject[key]

        if pruned > 0:
            logger.info(f"Pruned {pruned} entries older than {days} days")

    def add_entry(
        self,
        filename: str,
        date: str,
        sender_domain: str,
        original_subject: str,
        conversation_id: str,
        processing_status: str = "OPEN",
        has_attachments: bool = False,
        graph_conversation_id: Optional[str] = None,
        normalized_subject: Optional[str] = None,
    ):
        """Add an email to the index (both graph ID and subject lookups)."""
        entry = {
            "filename": filename,
            "date": date,
            "sender_domain": sender_domain,
            "original_subject": original_subject,
            "processing_status": processing_status,
            "has_attachments": has_attachments,
        }

        # Index by Graph conversationId (primary)
        if graph_conversation_id:
            if graph_conversation_id not in self.by_graph_id:
                self.by_graph_id[graph_conversation_id] = {
                    "conversation_id": conversation_id,
                    "entries": [],
                }
            group = self.by_graph_id[graph_conversation_id]
            existing_files = {e["filename"] for e in group["entries"]}
            if filename not in existing_files:
                group["entries"].append(entry)

        # Index by normalized subject (fallback)
        if normalized_subject:
            if normalized_subject not in self.by_subject:
                self.by_subject[normalized_subject] = {
                    "conversation_id": conversation_id,
                    "entries": [],
                }
                if graph_conversation_id:
                    self.by_subject[normalized_subject]["graph_id"] = graph_conversation_id
            group = self.by_subject[normalized_subject]
            existing_files = {e["filename"] for e in group["entries"]}
            if filename not in existing_files:
                group["entries"].append(entry)

    def find_conversation(
        self,
        graph_conversation_id: Optional[str],
        normalized_subject: str,
        sender_domain: str,
    ) -> Optional[tuple[str, list[dict]]]:
        """Find a matching conversation.

        Returns (conversation_id, entries) or None.
        Tries: graph_conversation_id → exact subject → Levenshtein subject.
        """
        # Step 1: Graph conversationId (most reliable — handles subject mutations, cross-domain)
        if graph_conversation_id and graph_conversation_id in self.by_graph_id:
            group = self.by_graph_id[graph_conversation_id]
            logger.debug(f"Graph ID match: {graph_conversation_id[:20]}...")
            return group["conversation_id"], group["entries"]

        # Step 2: Exact match on normalized subject
        if normalized_subject and normalized_subject in self.by_subject:
            group = self.by_subject[normalized_subject]
            return group["conversation_id"], group["entries"]

        # Step 3: Levenshtein fallback (same sender domain only)
        if not normalized_subject:
            return None

        best_match = None
        best_score = 0.0

        for idx_subject, group in self.by_subject.items():
            # Quick length check
            len_ratio = abs(len(normalized_subject) - len(idx_subject)) / max(
                len(normalized_subject), len(idx_subject), 1
            )
            if len_ratio > 0.5:
                continue

            # Check if at least one entry shares the sender domain
            domains = {e["sender_domain"] for e in group["entries"]}
            if sender_domain not in domains:
                continue

            score = levenshtein_similarity(normalized_subject, idx_subject)
            if score > best_score and score >= SUBJECT_SIMILARITY_THRESHOLD:
                best_score = score
                best_match = (group["conversation_id"], group["entries"])

        if best_match:
            logger.info(
                f"Levenshtein match ({best_score:.2f}): '{normalized_subject[:50]}'"
            )
        return best_match


def _sender_domain(email_dict: dict) -> str:
    """Extract sender domain from email dict."""
    addr = email_dict.get("from", {}).get("email", "")
    return addr.split("@")[1].lower() if "@" in addr else ""


def _email_filename(email_dict: dict) -> str:
    """Generate the standard filename for an email (same logic as output_formatter)."""
    import hashlib
    email_id = email_dict.get("id", "")
    email_hash = hashlib.md5(email_id.encode()).hexdigest()[:12]
    ts = email_dict.get("received_datetime", "").replace(":", "-").replace("T", "_").split(".")[0]
    return f"{ts}_{email_hash}.json"


def _generate_conversation_id(normalized_subject: str) -> str:
    """Generate a stable conversation ID from the normalized subject."""
    import hashlib
    return "conv_" + hashlib.md5(normalized_subject.encode()).hexdigest()[:12]


def _should_supersede_old(new_has_attachments: bool, existing_entries: list[dict]) -> bool:
    """Should the new email supersede (archive) the older ones?

    Returns True if old emails should be archived.
    Returns False if both should stay in inbox (old has attachments, new doesn't).

    Attachment rule: if an older email has attachments and the new one doesn't,
    both stay active — the old has the documents, the new has the latest context.
    """
    any_existing_has_attachments = any(e.get("has_attachments", False) for e in existing_entries)

    if not new_has_attachments and any_existing_has_attachments:
        return False  # Old has docs, new doesn't — keep both
    # New has docs, or tied, or neither has docs → newest supersedes old
    return True


def match_conversations(
    emails: list,
    email_folder: str,
) -> dict:
    """Match a batch of emails against the conversation index.

    Args:
        emails: List of Email objects (with .to_dict() method)
        email_folder: Path to the local email JSON folder (where index lives)

    Returns:
        Dict mapping email.id → conversation info:
        {
            "email_id": {
                "conversation_id": "conv_abc123",
                "is_latest": True,
                "is_chain": True,
                "position": 3,
                "related_emails": [
                    {"filename": "...", "date": "...", "subject": "...", "processing_status": "..."},
                    ...
                ],
                "supersedes": ["filename1.json", "filename2.json"]
            }
        }
    """
    index = ConversationIndex(email_folder)
    index.prune()

    results = {}

    # Sort emails oldest-first so we process in chronological order
    email_dicts = []
    for email in emails:
        d = email.to_dict()
        d["_email_obj"] = email  # keep reference
        email_dicts.append(d)
    email_dicts.sort(key=lambda d: d.get("received_datetime", ""))

    for email_dict in email_dicts:
        email_id = email_dict["id"]
        subject = email_dict.get("subject", "")
        norm_subj = normalize_subject(subject)
        sender_dom = _sender_domain(email_dict)
        filename = _email_filename(email_dict)
        date = email_dict.get("received_datetime", "")
        graph_conv_id = email_dict.get("graph_conversation_id")
        has_att = email_dict.get("has_attachments", False)

        if not norm_subj and not graph_conv_id:
            # No subject and no graph ID — no conversation matching possible
            conv_id = _generate_conversation_id(email_id)
            results[email_id] = {
                "conversation_id": conv_id,
                "is_latest": True,
                "is_chain": False,
                "position": 1,
                "related_emails": [],
                "supersedes": [],
            }
            index.add_entry(
                filename=filename, date=date, sender_domain=sender_dom,
                original_subject=subject, conversation_id=conv_id,
                has_attachments=has_att,
                graph_conversation_id=graph_conv_id,
                normalized_subject=norm_subj or email_id,
            )
            continue

        # Find existing conversation
        match = index.find_conversation(graph_conv_id, norm_subj, sender_dom)

        if match:
            conv_id, entries = match
            new_wins = _should_supersede_old(has_att, entries)

            related = [
                {
                    "filename": e["filename"],
                    "date": e["date"],
                    "subject": e["original_subject"],
                    "processing_status": e.get("processing_status", "OPEN"),
                }
                for e in entries
            ]

            if new_wins:
                # New email is the active one; supersede all prior
                supersedes = [e["filename"] for e in entries]
                results[email_id] = {
                    "conversation_id": conv_id,
                    "is_latest": True,
                    "is_chain": True,
                    "position": len(entries) + 1,
                    "related_emails": related,
                    "supersedes": supersedes,
                }
                # Mark previous batch emails as no longer latest
                for prev_entry in entries:
                    for other_dict in email_dicts:
                        if _email_filename(other_dict) == prev_entry["filename"]:
                            prev_id = other_dict["id"]
                            if prev_id in results:
                                results[prev_id]["is_latest"] = False
                            break
            else:
                # Old email has attachments, new doesn't — both stay in inbox
                # Neither supersedes the other (old has docs, new has latest context)
                results[email_id] = {
                    "conversation_id": conv_id,
                    "is_latest": True,
                    "is_chain": True,
                    "position": len(entries) + 1,
                    "related_emails": related,
                    "supersedes": [],
                }
                logger.info(
                    f"Attachment precedence: both stay active "
                    f"(old has attachments, new '{subject[:40]}' has latest context)"
                )

            # Add to index
            index.add_entry(
                filename=filename, date=date, sender_domain=sender_dom,
                original_subject=subject, conversation_id=conv_id,
                has_attachments=has_att,
                graph_conversation_id=graph_conv_id,
                normalized_subject=norm_subj,
            )

            logger.info(
                f"Conversation match: '{subject[:50]}' → conv {conv_id} "
                f"(pos {len(entries) + 1}, {len(entries)} prior, "
                f"new_wins={new_wins})"
            )
        else:
            # New conversation
            conv_id = _generate_conversation_id(norm_subj or email_id)
            results[email_id] = {
                "conversation_id": conv_id,
                "is_latest": True,
                "is_chain": False,
                "position": 1,
                "related_emails": [],
                "supersedes": [],
            }
            index.add_entry(
                filename=filename, date=date, sender_domain=sender_dom,
                original_subject=subject, conversation_id=conv_id,
                has_attachments=has_att,
                graph_conversation_id=graph_conv_id,
                normalized_subject=norm_subj,
            )

    # Save updated index
    index.save()

    # Mark is_chain on all emails that ended up in multi-email conversations
    conv_counts: dict[str, int] = {}
    for info in results.values():
        cid = info["conversation_id"]
        conv_counts[cid] = conv_counts.get(cid, 0) + 1
    # Also count index entries (from prior runs)
    for info in results.values():
        if info["position"] > 1:
            conv_counts[info["conversation_id"]] = max(
                conv_counts.get(info["conversation_id"], 0), info["position"]
            )
    for info in results.values():
        if conv_counts.get(info["conversation_id"], 1) > 1:
            info["is_chain"] = True

    # Summary
    multi = sum(1 for c in conv_counts.values() if c > 1)
    linked_to_existing = sum(1 for r in results.values() if r["position"] > 1)
    chains = sum(1 for r in results.values() if r["is_chain"])
    logger.info(
        f"Conversation matching: {len(results)} emails, "
        f"{len(conv_counts)} conversations, "
        f"{multi} multi-email, "
        f"{linked_to_existing} linked to existing, "
        f"{chains} in chains"
    )

    return results


def update_superseded_jsons(
    conversation_results: dict,
    email_folder: str,
) -> list[dict]:
    """Update older email JSONs to mark them as superseded.

    Sets is_latest_in_conversation=false, updates processing_status to
    ARCHIVE > SUPERSEDED, and adds conversation metadata.
    Only touches local JSON files.

    Returns:
        List of dicts with {"message_id": ..., "subject": ...} for old emails
        that were in active inbox states and need Outlook archiving.
    """
    folder = Path(email_folder)
    needs_outlook_archive = []

    for email_id, info in conversation_results.items():
        if not info["supersedes"]:
            continue

        for old_filename in info["supersedes"]:
            old_path = folder / old_filename
            if not old_path.exists():
                continue

            try:
                with open(old_path, "r", encoding="utf-8") as f:
                    data = json.load(f)

                # Only update if not already marked
                if data.get("is_latest_in_conversation") is False:
                    continue

                data["is_latest_in_conversation"] = False
                data["is_chain"] = True
                data["conversation_id"] = info["conversation_id"]
                data["conversation_position"] = data.get("conversation_position", info["position"] - 1)

                # Update processing_status: if still in an active state, mark superseded
                old_status = data.get("processing_status", "OPEN")
                if old_status in ("OPEN", "NEEDS REVIEW > VENDOR QUERY", "NEEDS REVIEW > NEW INVOICE"):
                    data["processing_status"] = "ARCHIVE > SUPERSEDED"
                    # This email was in inbox — needs Outlook archiving too
                    msg_id = data.get("id")
                    if msg_id:
                        needs_outlook_archive.append({
                            "message_id": msg_id,
                            "subject": data.get("subject", ""),
                        })

                with open(old_path, "w", encoding="utf-8") as f:
                    json.dump(data, f, indent=2, ensure_ascii=False)

                logger.info(f"Marked as superseded: {old_filename} (was: {old_status})")

            except Exception as e:
                logger.warning(f"Could not update {old_filename}: {e}")

    return needs_outlook_archive


def build_index_from_existing(email_folder: str) -> int:
    """Scan existing email JSONs and build the conversation index from scratch.

    Only reads subject, date, sender from each JSON (lightweight).
    Returns the number of entries indexed.
    """
    folder = Path(email_folder)
    if not folder.exists():
        logger.warning(f"Email folder not found: {email_folder}")
        return 0

    # Fresh index
    index = ConversationIndex(email_folder)
    index.by_graph_id = {}
    index.by_subject = {}

    cutoff = (datetime.now(timezone.utc) - timedelta(days=CONVERSATION_WINDOW_DAYS)).isoformat()

    json_files = sorted(folder.glob("*.json"))
    json_files = [f for f in json_files if f.name != "conversation_index.json"]

    count = 0
    for json_file in json_files:
        try:
            with open(json_file, "r", encoding="utf-8") as f:
                data = json.load(f)

            date = data.get("received_datetime", "")
            if date and date < cutoff:
                continue

            subject = data.get("subject", "")
            norm_subj = normalize_subject(subject)
            if not norm_subj:
                continue

            sender_email = data.get("from", {}).get("email", "")
            sender_dom = sender_email.split("@")[1].lower() if "@" in sender_email else ""
            graph_conv_id = data.get("graph_conversation_id")
            processing_status = data.get("processing_status", "OPEN")
            has_att = data.get("has_attachments", False)

            conv_id = _generate_conversation_id(norm_subj)

            index.add_entry(
                filename=json_file.name,
                date=date,
                sender_domain=sender_dom,
                original_subject=subject,
                conversation_id=conv_id,
                processing_status=processing_status,
                has_attachments=has_att,
                graph_conversation_id=graph_conv_id,
                normalized_subject=norm_subj,
            )
            count += 1

        except Exception as e:
            logger.warning(f"Error reading {json_file.name}: {e}")
            continue

    index.save()
    logger.info(f"Built conversation index: {count} emails indexed")
    return count
