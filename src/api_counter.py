"""Daily API call counter for Gemini models.

Tracks calls per model per day in a JSON file.
Used to manage the daily budget across normal runs and cleanup batches.
"""

import json
import logging
from datetime import date
from pathlib import Path

logger = logging.getLogger(__name__)

COUNTER_FILE = Path.home() / ".accounting_mailbox_reader" / "api_calls.json"
DAILY_BUDGET = 60  # Total across all models


def _load() -> dict:
    """Load counter data from disk."""
    if COUNTER_FILE.exists():
        try:
            with open(COUNTER_FILE, "r") as f:
                return json.load(f)
        except (json.JSONDecodeError, IOError):
            return {}
    return {}


def _save(data: dict):
    """Save counter data to disk."""
    COUNTER_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(COUNTER_FILE, "w") as f:
        json.dump(data, f, indent=2)


def increment(model: str) -> int:
    """Increment call count for model today. Returns new daily total."""
    today = str(date.today())
    data = _load()
    if today not in data:
        data[today] = {}
    data[today][model] = data[today].get(model, 0) + 1
    _save(data)
    total = sum(data[today].values())
    logger.debug(f"API call #{total} today ({model})")
    return total


def get_today_total() -> int:
    """Get total API calls made today across all models."""
    today = str(date.today())
    data = _load()
    if today not in data:
        return 0
    return sum(data[today].values())


def get_today_breakdown() -> dict:
    """Get per-model call counts for today."""
    today = str(date.today())
    data = _load()
    return data.get(today, {})


def get_remaining(budget: int = DAILY_BUDGET) -> int:
    """Get remaining API calls for today within budget."""
    return max(0, budget - get_today_total())
