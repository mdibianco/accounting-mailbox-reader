"""Teams notifications for the accounting mailbox agent."""

import logging
import os
from typing import Optional, Dict
from datetime import datetime

import requests

logger = logging.getLogger(__name__)


class TeamsNotifier:
    """Sends notifications to Teams via Incoming Webhook."""

    def __init__(self, webhook_url: Optional[str] = None):
        """
        Initialize Teams notifier.

        Args:
            webhook_url: Teams Incoming Webhook URL. If None, notifications are disabled.
        """
        self.webhook_url = webhook_url or os.environ.get("TEAMS_WEBHOOK_URL", "")
        self.enabled = bool(self.webhook_url)

        if self.enabled:
            logger.info("Teams notifier enabled")
        else:
            logger.debug("Teams notifier disabled (no webhook URL)")

    def notify_ven_inv_processed(
        self, email_dict: Dict, pass2_result: Dict
    ) -> bool:
        """
        Send notification for a processed VEN-INV email.

        Args:
            email_dict: Email dictionary
            pass2_result: Pass 2 processing result

        Returns:
            True if sent successfully, False otherwise
        """
        if not self.enabled:
            return False

        try:
            entity = pass2_result.get("planted_entity", {})
            entity_code = entity.get("code", "UNKNOWN")
            entity_name = entity.get("name", "Unknown")
            action = pass2_result.get("action_taken", "UNKNOWN")
            invoices_addr = pass2_result.get("invoices_address", "")

            from_email = email_dict.get("from", {}).get("email", "")
            from_name = email_dict.get("from", {}).get("name", from_email)
            subject = email_dict.get("subject", "")
            outlook_link = email_dict.get("outlook_link", "")

            # Build action description
            action_text = self._get_action_text(action, invoices_addr)
            color = self._get_action_color(action)

            # Adaptive Card
            card = {
                "type": "message",
                "attachments": [
                    {
                        "contentType": "application/vnd.microsoft.card.adaptive",
                        "contentUrl": None,
                        "content": {
                            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                            "type": "AdaptiveCard",
                            "version": "1.4",
                            "body": [
                                {
                                    "type": "Container",
                                    "style": "emphasis",
                                    "items": [
                                        {
                                            "type": "TextBlock",
                                            "text": f"🧾 Invoice Processed – {entity_code}",
                                            "weight": "bolder",
                                            "size": "large",
                                            "color": color,
                                        }
                                    ]
                                },
                                {
                                    "type": "Container",
                                    "items": [
                                        {
                                            "type": "FactSet",
                                            "facts": [
                                                {
                                                    "name": "Entity:",
                                                    "value": entity_name
                                                },
                                                {
                                                    "name": "From:",
                                                    "value": from_name
                                                },
                                                {
                                                    "name": "Subject:",
                                                    "value": subject[:60] + ("..." if len(subject) > 60 else "")
                                                },
                                                {
                                                    "name": "Action:",
                                                    "value": action_text
                                                },
                                                {
                                                    "name": "Invoices Address:",
                                                    "value": invoices_addr or "N/A"
                                                },
                                                {
                                                    "name": "Draft Reply:",
                                                    "value": "✓ Created" if pass2_result.get("draft_reply_created") else "—"
                                                }
                                            ]
                                        }
                                    ]
                                }
                            ],
                            "actions": [
                                {
                                    "type": "Action.OpenUrl",
                                    "title": "View in Outlook",
                                    "url": outlook_link
                                }
                            ]
                        }
                    }
                ]
            }

            return self._send_card(card)

        except Exception as e:
            logger.error(f"Error preparing VEN-INV notification: {e}")
            return False

    def notify_ven_inv_unknown_entity(self, email_dict: Dict) -> bool:
        """
        Send notification for a VEN-INV with unknown entity (needs review).

        Args:
            email_dict: Email dictionary

        Returns:
            True if sent successfully, False otherwise
        """
        if not self.enabled:
            return False

        try:
            from_email = email_dict.get("from", {}).get("email", "")
            from_name = email_dict.get("from", {}).get("name", from_email)
            subject = email_dict.get("subject", "")
            outlook_link = email_dict.get("outlook_link", "")

            card = {
                "type": "message",
                "attachments": [
                    {
                        "contentType": "application/vnd.microsoft.card.adaptive",
                        "contentUrl": None,
                        "content": {
                            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                            "type": "AdaptiveCard",
                            "version": "1.4",
                            "body": [
                                {
                                    "type": "Container",
                                    "style": "emphasis",
                                    "items": [
                                        {
                                            "type": "TextBlock",
                                            "text": "⚠️ Invoice – Unknown Entity",
                                            "weight": "bolder",
                                            "size": "large",
                                            "color": "warning",
                                        }
                                    ]
                                },
                                {
                                    "type": "Container",
                                    "items": [
                                        {
                                            "type": "TextBlock",
                                            "text": "This invoice could not be matched to a Planted entity. Please review and determine the correct invoices address.",
                                            "wrap": True
                                        },
                                        {
                                            "type": "FactSet",
                                            "facts": [
                                                {
                                                    "name": "From:",
                                                    "value": from_name
                                                },
                                                {
                                                    "name": "Subject:",
                                                    "value": subject[:60] + ("..." if len(subject) > 60 else "")
                                                }
                                            ]
                                        }
                                    ]
                                }
                            ],
                            "actions": [
                                {
                                    "type": "Action.OpenUrl",
                                    "title": "Review in Outlook",
                                    "url": outlook_link
                                }
                            ]
                        }
                    }
                ]
            }

            return self._send_card(card)

        except Exception as e:
            logger.error(f"Error preparing unknown entity notification: {e}")
            return False

    def notify_run_summary(self, processed_count: int, unknown_count: int) -> bool:
        """
        Send a summary notification at the end of a run.

        Args:
            processed_count: Number of VEN-INV emails successfully processed
            unknown_count: Number of VEN-INV emails with unknown entity

        Returns:
            True if sent successfully, False otherwise
        """
        if not self.enabled:
            return False

        try:
            card = {
                "type": "message",
                "attachments": [
                    {
                        "contentType": "application/vnd.microsoft.card.adaptive",
                        "contentUrl": None,
                        "content": {
                            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                            "type": "AdaptiveCard",
                            "version": "1.4",
                            "body": [
                                {
                                    "type": "Container",
                                    "style": "emphasis",
                                    "items": [
                                        {
                                            "type": "TextBlock",
                                            "text": "✓ Accounting Agent Run Complete",
                                            "weight": "bolder",
                                            "size": "large",
                                            "color": "good",
                                        }
                                    ]
                                },
                                {
                                    "type": "Container",
                                    "items": [
                                        {
                                            "type": "FactSet",
                                            "facts": [
                                                {
                                                    "name": "VEN-INV Processed:",
                                                    "value": str(processed_count)
                                                },
                                                {
                                                    "name": "Unknown Entities:",
                                                    "value": str(unknown_count)
                                                },
                                                {
                                                    "name": "Timestamp:",
                                                    "value": datetime.utcnow().isoformat() + "Z"
                                                }
                                            ]
                                        }
                                    ]
                                }
                            ]
                        }
                    }
                ]
            }

            return self._send_card(card)

        except Exception as e:
            logger.error(f"Error preparing summary notification: {e}")
            return False

    def _send_card(self, card: Dict) -> bool:
        """
        Send an Adaptive Card to Teams.

        Args:
            card: Adaptive Card payload

        Returns:
            True if sent successfully, False otherwise
        """
        if not self.enabled:
            return False

        try:
            response = requests.post(
                self.webhook_url,
                json=card,
                timeout=10
            )
            response.raise_for_status()
            logger.debug("Teams notification sent successfully")
            return True
        except requests.exceptions.RequestException as e:
            logger.error(f"Failed to send Teams notification: {e}")
            return False

    @staticmethod
    def _get_action_text(action: str, invoices_addr: str) -> str:
        """Get human-readable action text."""
        if action == "FORWARDED":
            return f"Forwarded to {invoices_addr}"
        elif action == "ALREADY_IN_CC":
            return "Already in CC – no action needed"
        elif action == "FOUND_IN_MAILBOX":
            return f"Already received at {invoices_addr}"
        elif action == "UNKNOWN_ENTITY":
            return "Unknown entity – manual review needed"
        else:
            return action

    @staticmethod
    def _get_action_color(action: str) -> str:
        """Get card color based on action."""
        if action == "UNKNOWN_ENTITY":
            return "warning"
        elif action in ("FORWARDED", "FOUND_IN_MAILBOX"):
            return "accent"
        else:
            return "good"
