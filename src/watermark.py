"""Watermark tracking — remembers the last processed email datetime."""

import json
import logging
from datetime import datetime, timezone
from pathlib import Path

logger = logging.getLogger(__name__)

WATERMARK_FILE = Path.home() / ".accounting_mailbox_reader" / "watermark.json"

# Seed date: everything before this is considered already processed
SEED_DATE = "2026-02-27T00:00:00Z"


def get_watermark() -> str:
    """Get the last processed datetime. Returns seed date if no watermark exists."""
    if WATERMARK_FILE.exists():
        try:
            with open(WATERMARK_FILE, "r") as f:
                data = json.load(f)
            return data.get("last_processed_datetime", SEED_DATE)
        except (json.JSONDecodeError, KeyError):
            logger.warning(f"Corrupt watermark file, using seed date: {SEED_DATE}")
    return SEED_DATE


def update_watermark(last_datetime: str):
    """Update the watermark to the given datetime string."""
    WATERMARK_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(WATERMARK_FILE, "w") as f:
        json.dump({
            "last_processed_datetime": last_datetime,
            "updated_at": datetime.now(timezone.utc).isoformat(),
        }, f, indent=2)
    logger.info(f"Watermark updated to: {last_datetime}")
