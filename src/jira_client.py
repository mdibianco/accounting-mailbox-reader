"""Jira integration for creating tickets from classified emails."""

import json
import logging
from base64 import b64encode
from typing import Optional

import requests

from .config import config

logger = logging.getLogger(__name__)

# Trigger -> assignee mapping
ASSIGNEE_MAP = {
    "PRIO_HIGHEST": "712020:f3cf8d28-7d95-4cc8-9769-155bf6451608",   # Jalal Hanai
    "VEN-FOLLOWUP": "712020:cd6de4a5-13ce-4e0f-acbd-326d21005431",   # Dimitria Milanova
    "OTHER":        "6396e8f13c9bcd363976d34e",                       # Matthias Di Bianco
}

TRIGGER_REASONS = {
    "PRIO_HIGHEST": "Highest priority flag - requires immediate attention",
    "VEN-FOLLOWUP": "Vendor query requiring active accounting response",
    "OTHER": "Unclassified high-priority email requiring manual review",
}

JIRA_BASE_URL = "https://eatplanted.atlassian.net"
JIRA_PROJECT_KEY = "FH20"
JIRA_ISSUE_TYPE_NAME = "General request"


class JiraClient:
    """Creates Jira tickets in FH20 Finance Helpdesk. No-ops if not configured."""

    def __init__(self):
        self.user_email = getattr(config, "jira_user_email", None) or ""
        self.api_token = getattr(config, "jira_api_token", None) or ""
        self.enabled = bool(self.user_email.strip() and self.api_token.strip())

        if not self.enabled:
            logger.info("Jira integration disabled (JIRA_USER_EMAIL / JIRA_API_TOKEN not set)")
            return

        creds = b64encode(f"{self.user_email}:{self.api_token}".encode()).decode()
        self.headers = {
            "Authorization": f"Basic {creds}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        }
        self.base_url = JIRA_BASE_URL
        logger.info("Jira integration enabled for %s", self.user_email)

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def should_create_ticket(self, email_dict: dict) -> bool:
        """Check chain rules and existing ticket before creating."""
        if not self.enabled:
            return False
        if email_dict.get("jira_issue_key"):
            return False
        if not email_dict.get("is_latest_in_conversation", True):
            return False
        status = email_dict.get("processing_status", "")
        if status.startswith("ARCHIVE"):
            return False
        return True

    def find_existing_ticket(self, email_id: str) -> Optional[str]:
        """Search FH20 for a ticket that already references this email ID."""
        if not self.enabled or not email_id:
            return None
        jql = f'project = {JIRA_PROJECT_KEY} AND description ~ "{email_id}"'
        params = {"jql": jql, "fields": "key", "maxResults": 1}
        try:
            resp = requests.get(
                f"{self.base_url}/rest/api/3/search",
                headers=self.headers,
                params=params,
                timeout=15,
            )
            resp.raise_for_status()
            issues = resp.json().get("issues", [])
            if issues:
                key = issues[0]["key"]
                logger.info("Found existing Jira ticket %s for email %s", key, email_id[:20])
                return key
        except Exception as e:
            logger.warning("Jira search failed: %s", e)
        return None

    def create_ticket(self, email_dict: dict, trigger: str) -> Optional[str]:
        """Create a Jira ticket in FH20 and return the issue key."""
        if not self.enabled:
            return None

        subject = email_dict.get("subject", "No subject")
        summary = f"MAIL: {subject}"[:255]
        description_adf = self._build_description(email_dict, trigger)
        assignee_id = ASSIGNEE_MAP.get(trigger)

        payload = {
            "fields": {
                "project": {"key": JIRA_PROJECT_KEY},
                "issuetype": {"name": JIRA_ISSUE_TYPE_NAME},
                "summary": summary,
                "description": description_adf,
            }
        }
        if assignee_id:
            payload["fields"]["assignee"] = {"accountId": assignee_id}

        try:
            resp = requests.post(
                f"{self.base_url}/rest/api/3/issue",
                headers=self.headers,
                json=payload,
                timeout=30,
            )
            resp.raise_for_status()
            key = resp.json().get("key", "")
            logger.info("Created Jira ticket %s (trigger=%s, subject=%s)", key, trigger, subject[:50])
            return key
        except requests.exceptions.HTTPError as e:
            logger.error("Jira create failed (%s): %s", e.response.status_code, e.response.text[:500])
        except Exception as e:
            logger.error("Jira create failed: %s", e)
        return None

    # ------------------------------------------------------------------
    # Internal
    # ------------------------------------------------------------------

    def _build_description(self, email_dict: dict, trigger: str) -> dict:
        """Build Atlassian Document Format (ADF) description."""
        reason = TRIGGER_REASONS.get(trigger, "Flagged by mailbox agent")
        outlook_link = email_dict.get("outlook_link", "")
        from_info = email_dict.get("from", {})
        from_name = from_info.get("name", "")
        from_email = from_info.get("email", "")
        received = email_dict.get("received_datetime", "")
        priority = email_dict.get("classification", {}).get("priority", "")
        entity_code = ""
        entity_name = ""
        p2 = email_dict.get("pass2_results") or {}
        pe = p2.get("planted_entity") or {}
        entity_code = pe.get("code", "")
        entity_name = pe.get("name", "")

        json_dump = json.dumps(email_dict, indent=2, ensure_ascii=False, default=str)
        # Truncate JSON to avoid hitting Jira's 32KB description limit
        if len(json_dump) > 25000:
            json_dump = json_dump[:25000] + "\n\n... [truncated]"

        # Build ADF (Atlassian Document Format)
        content = []

        # Trigger heading
        content.append(self._adf_heading("Trigger", 2))
        content.append(self._adf_paragraph(f"Created by accounting mailbox agent: {trigger}"))
        content.append(self._adf_paragraph(reason))

        # Email details heading
        content.append(self._adf_heading("Email", 2))
        content.append(self._adf_paragraph(f"From: {from_name} ({from_email})"))
        content.append(self._adf_paragraph(f"Date: {received}"))
        if entity_code:
            content.append(self._adf_paragraph(f"Entity: {entity_code} - {entity_name}"))
        content.append(self._adf_paragraph(f"Priority: {priority}"))
        if outlook_link:
            content.append({
                "type": "paragraph",
                "content": [{
                    "type": "text",
                    "text": "Open in Outlook",
                    "marks": [{"type": "link", "attrs": {"href": outlook_link}}],
                }],
            })

        # Raw data
        content.append(self._adf_heading("Raw Data", 2))
        content.append({
            "type": "codeBlock",
            "attrs": {"language": "json"},
            "content": [{"type": "text", "text": json_dump}],
        })

        return {
            "version": 1,
            "type": "doc",
            "content": content,
        }

    @staticmethod
    def _adf_heading(text: str, level: int = 2) -> dict:
        return {
            "type": "heading",
            "attrs": {"level": level},
            "content": [{"type": "text", "text": text}],
        }

    @staticmethod
    def _adf_paragraph(text: str) -> dict:
        return {
            "type": "paragraph",
            "content": [{"type": "text", "text": text}],
        }
