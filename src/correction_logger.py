"""Scan ARCHIVE/WRONG_CLASSIFICATION folder and log corrections to YAML."""

import hashlib
import json
import logging
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional

import yaml

from .graph_client import GraphAPIClient
from .keyword_triage import KeywordTriage

logger = logging.getLogger(__name__)

CORRECTIONS_FILE = Path(__file__).parent.parent / "config" / "corrections.yaml"
WRONG_FOLDER = "ARCHIVE/WRONG_CLASSIFICATION"

# Valid category IDs that can appear as Outlook categories
VALID_CATEGORIES = set(KeywordTriage.CATEGORY_NAMES.keys())


def scan_corrections(
    graph: GraphAPIClient,
    mailbox: str,
    local_folder: Optional[str] = None,
    dry_run: bool = False,
) -> int:
    """
    Scan ARCHIVE/WRONG_CLASSIFICATION for corrected emails and log them.

    The user drags misclassified emails to this folder and sets the Outlook
    category to the correct classification. We read the category, find the
    original JSON, and append a correction entry to config/corrections.yaml.

    Emails stay in WRONG_CLASSIFICATION as a permanent edge-case reference.

    Returns number of new corrections logged.
    """
    folder_id = graph.get_folder_id(mailbox, WRONG_FOLDER)
    if not folder_id:
        logger.debug("WRONG_CLASSIFICATION folder not found — nothing to scan")
        return 0

    messages = graph.get_folder_messages(
        mailbox, folder_id, max_results=500, days_back=365
    )
    if not messages:
        return 0

    # Load existing corrections to avoid duplicates
    existing = _load_corrections()
    logged_ids = {c["email_id"] for c in existing}

    new_count = 0
    for msg in messages:
        msg_id = msg["id"]
        if msg_id in logged_ids:
            continue

        # Extract correct category from Outlook categories
        outlook_cats = msg.get("categories", [])
        correct_category = _extract_category(outlook_cats)
        if not correct_category:
            logger.warning(
                f"No valid category on corrected email: {msg.get('subject', '?')[:60]} "
                f"(categories: {outlook_cats})"
            )
            continue

        # Find original JSON to get the original classification
        original_category = None
        original_priority = None
        if local_folder:
            json_data = _find_email_json(msg_id, local_folder)
            if json_data:
                orig_cls = json_data.get("classification", {})
                original_category = orig_cls.get("primary_category", {}).get("id")
                original_priority = orig_cls.get("priority")

        # Skip if the correction matches the original (user may not have changed it)
        if original_category and original_category == correct_category:
            logger.info(
                f"Skipping — correction matches original ({correct_category}): "
                f"{msg.get('subject', '?')[:60]}"
            )
            continue

        # Build correction entry
        from_data = msg.get("from", {}).get("emailAddress", {})
        to_list = [
            r.get("emailAddress", {}).get("address", "")
            for r in msg.get("toRecipients", [])
        ]
        cc_list = [
            r.get("emailAddress", {}).get("address", "")
            for r in msg.get("ccRecipients", [])
        ]

        entry = {
            "email_id": msg_id,
            "date": msg.get("receivedDateTime", "")[:10],
            "subject": msg.get("subject", ""),
            "from": from_data.get("address", ""),
            "to": to_list,
            "cc": cc_list,
            "original_category": original_category,
            "original_priority": original_priority,
            "correct_category": correct_category,
            "reasoning": "",  # User fills this in
            "logged_at": datetime.now(timezone.utc).isoformat(),
        }

        if not dry_run:
            existing.append(entry)
            new_count += 1
            logger.info(
                f"Logged correction: {original_category} -> {correct_category}: "
                f"{msg.get('subject', '?')[:60]}"
            )
        else:
            logger.info(
                f"(dry run) Would log correction: {original_category} -> {correct_category}: "
                f"{msg.get('subject', '?')[:60]}"
            )
            new_count += 1

    if not dry_run and new_count > 0:
        _save_corrections(existing)

    return new_count


def _extract_category(outlook_categories: list) -> Optional[str]:
    """Extract the first valid category ID from Outlook categories."""
    for cat in outlook_categories:
        cat_upper = cat.upper().replace(" ", "_")
        if cat_upper in VALID_CATEGORIES:
            return cat_upper
        # Also check original case (categories are set as-is)
        if cat in VALID_CATEGORIES:
            return cat
    return None


def _find_email_json(email_id: str, local_folder: str) -> Optional[dict]:
    """Find and read the JSON file for an email by its ID hash."""
    email_hash = hashlib.md5(email_id.encode()).hexdigest()[:12]
    folder = Path(local_folder)

    matches = list(folder.glob(f"*_{email_hash}.json"))
    if not matches:
        return None

    try:
        with open(matches[0], "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        logger.warning(f"Could not read JSON for {email_hash}: {e}")
        return None


def _load_corrections() -> list:
    """Load existing corrections from YAML file."""
    if not CORRECTIONS_FILE.exists():
        return []
    try:
        with open(CORRECTIONS_FILE, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f)
        return data if isinstance(data, list) else []
    except Exception as e:
        logger.warning(f"Could not read corrections file: {e}")
        return []


def get_pending_corrections_count() -> int:
    """Count total corrections that haven't been processed into rule/prompt updates."""
    corrections = _load_corrections()
    return len(corrections)


def _save_corrections(corrections: list):
    """Save corrections to YAML file."""
    CORRECTIONS_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(CORRECTIONS_FILE, "w", encoding="utf-8") as f:
        yaml.dump(
            corrections,
            f,
            default_flow_style=False,
            allow_unicode=True,
            sort_keys=False,
        )
    logger.info(f"Saved {len(corrections)} corrections to {CORRECTIONS_FILE}")
