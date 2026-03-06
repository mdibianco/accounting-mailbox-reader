"""Daily statistics tracking and summary generation."""

import json
import logging
from datetime import datetime, timezone, timedelta
from pathlib import Path
from typing import Dict, List, Optional

logger = logging.getLogger(__name__)

LAST_SUMMARY_FILE = Path.home() / ".accounting_mailbox_reader" / "last_summary_sent.json"


class DailyStats:
    """Track and manage daily processing statistics."""

    STATS_DIR = Path.home() / ".accounting_mailbox_reader"

    @classmethod
    def _get_stats_file(cls, date: Optional[str] = None) -> Path:
        """Get the stats file path for a given date (YYYY-MM-DD)."""
        if date is None:
            date = datetime.utcnow().strftime("%Y-%m-%d")
        cls.STATS_DIR.mkdir(exist_ok=True)
        return cls.STATS_DIR / f"daily_stats_{date}.json"

    @classmethod
    def _load_stats(cls, date: Optional[str] = None) -> Dict:
        """Load existing stats or create new ones."""
        stats_file = cls._get_stats_file(date)
        if stats_file.exists():
            try:
                with open(stats_file, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception as e:
                logger.error(f"Failed to load stats: {e}")
                return cls._init_stats()
        return cls._init_stats()

    @classmethod
    def _init_stats(cls) -> Dict:
        """Initialize empty stats structure."""
        return {
            "date": datetime.utcnow().strftime("%Y-%m-%d"),
            "scheduled_runs": {
                "total_emails_processed": 0,
                "categories_by_keywords": 0,
                "categories_by_llm": 0,
                "total_ven_rem_analyzed": 0,
                "total_ven_followup_analyzed": 0,
                "total_ven_inv_processed": 0,
                "total_emails_archived": 0,
                "total_human_completed_moved": 0,
                "total_jsons_saved": 0,
                "llm_calls_by_model": {},
                "run_count": 0,
            },
            "cleanup_run": {
                "emails_in_folder": 0,
                "emails_classified": 0,
                "emails_with_pass2": 0,
                "emails_moved_to_archive": 0,
                "api_calls_used": 0,
            },
        }

    @classmethod
    def _save_stats(cls, stats: Dict, date: Optional[str] = None) -> None:
        """Save stats to file."""
        stats_file = cls._get_stats_file(date)
        try:
            with open(stats_file, "w", encoding="utf-8") as f:
                json.dump(stats, f, indent=2, ensure_ascii=False)
        except Exception as e:
            logger.error(f"Failed to save stats: {e}")

    @classmethod
    def record_process_run(
        cls,
        emails_processed: int,
        categories_by_keywords: int,
        categories_by_llm: int,
        ven_rem_analyzed: int,
        ven_followup_analyzed: int,
        ven_inv_processed: int,
        emails_archived: int,
        human_completed_moved: int,
        jsons_saved: int,
        llm_calls_by_model: Dict = None,
    ) -> None:
        """Record stats from a scheduled process run."""
        if llm_calls_by_model is None:
            llm_calls_by_model = {}

        stats = cls._load_stats()
        stats["scheduled_runs"]["total_emails_processed"] += emails_processed
        stats["scheduled_runs"]["categories_by_keywords"] += categories_by_keywords
        stats["scheduled_runs"]["categories_by_llm"] += categories_by_llm
        stats["scheduled_runs"]["total_ven_rem_analyzed"] += ven_rem_analyzed
        stats["scheduled_runs"]["total_ven_followup_analyzed"] += ven_followup_analyzed
        stats["scheduled_runs"]["total_ven_inv_processed"] += ven_inv_processed
        stats["scheduled_runs"]["total_emails_archived"] += emails_archived
        stats["scheduled_runs"]["total_human_completed_moved"] += human_completed_moved
        stats["scheduled_runs"]["total_jsons_saved"] += jsons_saved

        # Accumulate LLM calls by model
        for model, count in llm_calls_by_model.items():
            if model not in stats["scheduled_runs"]["llm_calls_by_model"]:
                stats["scheduled_runs"]["llm_calls_by_model"][model] = 0
            stats["scheduled_runs"]["llm_calls_by_model"][model] += count

        stats["scheduled_runs"]["run_count"] += 1
        cls._save_stats(stats)

        # Also save process run record to emails folder for Power Query access
        cls._save_process_run_log(
            emails_processed, categories_by_keywords, categories_by_llm,
            ven_rem_analyzed, ven_followup_analyzed, ven_inv_processed,
            emails_archived, human_completed_moved, jsons_saved,
            llm_calls_by_model
        )

        logger.info("Recorded process run stats")

    @classmethod
    def record_cleanup_run(
        cls,
        emails_in_folder: int,
        emails_classified: int,
        emails_with_pass2: int,
        emails_moved_to_archive: int,
        api_calls_used: int,
    ) -> None:
        """Record stats from the cleanup run."""
        stats = cls._load_stats()
        stats["cleanup_run"]["emails_in_folder"] = emails_in_folder
        stats["cleanup_run"]["emails_classified"] = emails_classified
        stats["cleanup_run"]["emails_with_pass2"] = emails_with_pass2
        stats["cleanup_run"]["emails_moved_to_archive"] = emails_moved_to_archive
        stats["cleanup_run"]["api_calls_used"] = api_calls_used
        cls._save_stats(stats)

        # Also save cleanup run record to emails folder for Power Query access
        cls._save_cleanup_run_log(
            emails_in_folder, emails_classified, emails_with_pass2,
            emails_moved_to_archive, api_calls_used
        )

        logger.info("Recorded cleanup run stats")

    @classmethod
    def get_daily_summary(cls) -> Optional[Dict]:
        """Get today's accumulated statistics."""
        return cls._load_stats()

    @classmethod
    def reset_daily_stats(cls) -> None:
        """Reset stats for the next day."""
        stats = cls._init_stats()
        cls._save_stats(stats)
        logger.info("Daily stats reset")

    @classmethod
    def get_last_summary_sent(cls) -> Optional[str]:
        """Get timestamp of last summary email sent."""
        if LAST_SUMMARY_FILE.exists():
            try:
                with open(LAST_SUMMARY_FILE, "r") as f:
                    data = json.load(f)
                return data.get("last_sent")
            except Exception:
                pass
        return None

    @classmethod
    def set_last_summary_sent(cls) -> None:
        """Record that a summary email was just sent."""
        LAST_SUMMARY_FILE.parent.mkdir(exist_ok=True)
        with open(LAST_SUMMARY_FILE, "w") as f:
            json.dump({"last_sent": datetime.now(timezone.utc).isoformat()}, f)

    @classmethod
    def aggregate_runs_since_last_summary(cls) -> Dict:
        """Aggregate process_runs.json entries since last summary was sent.

        Returns a dict with totals suitable for the summary email.
        """
        from src.config import config

        last_sent = cls.get_last_summary_sent()
        if last_sent:
            cutoff = datetime.fromisoformat(last_sent.replace("Z", "+00:00"))
        else:
            # Default: last 24 hours
            cutoff = datetime.now(timezone.utc) - timedelta(hours=24)

        result = {
            "since": cutoff.isoformat(),
            "run_count": 0,
            "total_emails_processed": 0,
            "categories_by_keywords": 0,
            "categories_by_llm": 0,
            "total_ven_rem_analyzed": 0,
            "total_ven_followup_analyzed": 0,
            "total_ven_inv_processed": 0,
            "total_emails_archived": 0,
            "total_human_completed_moved": 0,
            "total_jsons_saved": 0,
            "llm_calls_by_model": {},
        }

        if not config.local_folder_path:
            return result

        process_runs_file = Path(config.local_folder_path) / "process_runs.json"
        if not process_runs_file.exists():
            return result

        try:
            with open(process_runs_file, "r", encoding="utf-8") as f:
                runs = json.load(f)
        except Exception:
            return result

        for run in runs:
            ts = run.get("timestamp", "")
            try:
                run_dt = datetime.fromisoformat(ts.replace("Z", "+00:00"))
            except (ValueError, AttributeError):
                continue
            if run_dt <= cutoff:
                continue

            result["run_count"] += 1
            result["total_emails_processed"] += run.get("emails_processed", 0)
            result["categories_by_keywords"] += run.get("categories_by_keywords", 0)
            result["categories_by_llm"] += run.get("categories_by_llm", 0)
            result["total_ven_rem_analyzed"] += run.get("ven_rem_analyzed", 0)
            result["total_ven_followup_analyzed"] += run.get("ven_followup_analyzed", 0)
            result["total_ven_inv_processed"] += run.get("ven_inv_processed", 0)
            result["total_emails_archived"] += run.get("emails_archived", 0)
            result["total_human_completed_moved"] += run.get("human_completed_moved", 0)
            result["total_jsons_saved"] += run.get("jsons_saved", 0)
            for model, count in run.get("llm_calls_by_model", {}).items():
                result["llm_calls_by_model"][model] = result["llm_calls_by_model"].get(model, 0) + count

        return result

    @classmethod
    def _save_process_run_log(
        cls,
        emails_processed: int,
        categories_by_keywords: int,
        categories_by_llm: int,
        ven_rem_analyzed: int,
        ven_followup_analyzed: int,
        ven_inv_processed: int,
        emails_archived: int,
        human_completed_moved: int,
        jsons_saved: int,
        llm_calls_by_model: Dict,
    ) -> None:
        """Save individual process run log to emails folder for Power Query."""
        try:
            from src.config import config

            if not config.local_folder_path:
                return

            # Create process_runs JSON if it doesn't exist
            process_runs_file = Path(config.local_folder_path) / "process_runs.json"
            if process_runs_file.exists():
                with open(process_runs_file, "r", encoding="utf-8") as f:
                    runs = json.load(f)
            else:
                runs = []

            # Add new run record
            runs.append({
                "timestamp": datetime.utcnow().isoformat() + "Z",
                "type": "scheduled_run",
                "emails_processed": emails_processed,
                "categories_by_keywords": categories_by_keywords,
                "categories_by_llm": categories_by_llm,
                "ven_rem_analyzed": ven_rem_analyzed,
                "ven_followup_analyzed": ven_followup_analyzed,
                "ven_inv_processed": ven_inv_processed,
                "emails_archived": emails_archived,
                "human_completed_moved": human_completed_moved,
                "jsons_saved": jsons_saved,
                "llm_calls_by_model": llm_calls_by_model,
            })

            # Keep only last 365 days of runs (8760 hourly runs)
            cutoff_date = datetime.now(timezone.utc) - timedelta(days=365)
            runs = [r for r in runs if datetime.fromisoformat(r["timestamp"].replace("Z", "+00:00")) > cutoff_date]

            with open(process_runs_file, "w", encoding="utf-8") as f:
                json.dump(runs, f, indent=2, ensure_ascii=False, default=str)

            logger.info(f"Logged process run to {process_runs_file.name}")
        except Exception as e:
            logger.error(f"Failed to save process run log: {e}")

    @classmethod
    def _save_cleanup_run_log(
        cls,
        emails_in_folder: int,
        emails_classified: int,
        emails_with_pass2: int,
        emails_moved_to_archive: int,
        api_calls_used: int,
    ) -> None:
        """Save cleanup run log to emails folder for Power Query."""
        try:
            from src.config import config

            if not config.local_folder_path:
                return

            # Create cleanup_runs JSON if it doesn't exist
            cleanup_runs_file = Path(config.local_folder_path) / "cleanup_runs.json"
            if cleanup_runs_file.exists():
                with open(cleanup_runs_file, "r", encoding="utf-8") as f:
                    runs = json.load(f)
            else:
                runs = []

            # Add new cleanup run record
            runs.append({
                "timestamp": datetime.utcnow().isoformat() + "Z",
                "type": "cleanup_run",
                "emails_in_folder": emails_in_folder,
                "emails_classified": emails_classified,
                "emails_with_pass2": emails_with_pass2,
                "emails_moved_to_archive": emails_moved_to_archive,
                "api_calls_used": api_calls_used,
            })

            # Keep only last 90 days of cleanup runs
            cutoff_date = datetime.now(timezone.utc) - timedelta(days=90)
            runs = [r for r in runs if datetime.fromisoformat(r["timestamp"].replace("Z", "+00:00")) > cutoff_date]

            with open(cleanup_runs_file, "w", encoding="utf-8") as f:
                json.dump(runs, f, indent=2, ensure_ascii=False, default=str)

            logger.info(f"Logged cleanup run to {cleanup_runs_file.name}")
        except Exception as e:
            logger.error(f"Failed to save cleanup run log: {e}")
