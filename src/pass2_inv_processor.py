"""Pass 2 processor for VEN-INV (vendor invoice) emails."""

import json
import logging
import re
from datetime import datetime, timedelta
from typing import Dict, Optional, List

from .config import config

logger = logging.getLogger(__name__)


class Pass2InvProcessor:
    """Processes VEN-INV emails: detects entity, checks CC, forwards if needed, creates draft reply."""

    def __init__(self):
        """Initialize VEN-INV processor."""
        self.entity_names = config.entity_names  # {code: name} mapping

    def process_email(
        self, email_dict: Dict, graph_client=None, dry_run: bool = False  # noqa: ARG002 — kept for future re-enable
    ) -> Optional[Dict]:
        """
        Process a VEN-INV email: entity detection → CC check → forward/draft reply.

        Args:
            email_dict: Email dict from Email.to_dict()
            graph_client: GraphAPIClient instance
            dry_run: If True, don't actually forward or create drafts

        Returns:
            Dictionary with pass2_results, or None on error
        """
        try:
            # Step 1: Detect entity from email content (fuzzy text search)
            entity_code, entity_name = self._detect_entity(email_dict)

            if entity_code == "UNKNOWN":
                logger.warning(f"VEN-INV: Unknown entity in email {email_dict.get('subject')}")
                return {
                    "pass2_timestamp": datetime.utcnow().isoformat() + "Z",
                    "pass2_model": "rule-based",
                    "planted_entity": {
                        "code": "UNKNOWN",
                        "name": "Unknown"
                    },
                    "invoices_address": None,
                    "action_taken": "UNKNOWN_ENTITY",
                    "forwarded_to": None,
                    "draft_reply_id": None,
                    "draft_reply_created": False,
                    "invoices": [],
                    "urgency_level": None,
                }

            # Step 2: Get invoices address for this entity
            invoices_addr = self._get_invoices_address(entity_code)
            logger.info(f"VEN-INV: Entity={entity_code}, invoices_addr={invoices_addr}")

            # Step 3: Check if already in CC
            if self._is_in_cc(email_dict, invoices_addr):
                logger.info(f"VEN-INV: {invoices_addr} already in CC - no action needed")
                return {
                    "pass2_timestamp": datetime.utcnow().isoformat() + "Z",
                    "pass2_model": "rule-based",
                    "planted_entity": {
                        "code": entity_code,
                        "name": entity_name
                    },
                    "invoices_address": invoices_addr,
                    "action_taken": "ALREADY_IN_CC",
                    "forwarded_to": None,
                    "draft_reply_id": None,
                    "draft_reply_created": False,
                    "invoices": [],
                    "urgency_level": None,
                }

            # Step 4: Check if email reached the invoices mailbox (CH1 only for now)
            found_in_mailbox = False
            if entity_code == "CH1":
                found_in_mailbox = self._check_invoices_mailbox(
                    email_dict, graph_client
                )
                if found_in_mailbox:
                    logger.info(f"VEN-INV: Email found in {invoices_addr} inbox")

            if found_in_mailbox:
                action = "FOUND_IN_MAILBOX"
                forwarded_to = None
            else:
                # Step 5: Forward the email — DISABLED, will be picked up later
                forwarded_to = None
                action = "FORWARDING_DISABLED"
                logger.info(f"VEN-INV: [FORWARDING DISABLED] Would forward to {invoices_addr}")

            # Step 6: Draft reply — DISABLED, VEN-INV stays in inbox for manual review
            draft_reply_id = None
            draft_created = False
            logger.info("VEN-INV: [DRAFT REPLY DISABLED] Would create draft reply")

            return {
                "pass2_timestamp": datetime.utcnow().isoformat() + "Z",
                "pass2_model": "rule-based",
                "planted_entity": {
                    "code": entity_code,
                    "name": entity_name
                },
                "invoices_address": invoices_addr,
                "action_taken": action,
                "forwarded_to": forwarded_to,
                "draft_reply_id": draft_reply_id,
                "draft_reply_created": draft_created,
                "invoices": [],
                "urgency_level": None,
            }

        except Exception as e:
            logger.error(f"VEN-INV processing failed: {e}", exc_info=True)
            return None

    def _detect_entity(self, email_dict: Dict) -> tuple:
        """
        Detect entity from email content (fuzzy text search).

        Returns:
            (entity_code, entity_name) tuple, or ("UNKNOWN", "Unknown")
        """
        # Build search text from subject, body, and attachment text
        search_parts = [
            email_dict.get("subject", ""),
            email_dict.get("body", ""),
        ]

        # Add attachment text
        for att in email_dict.get("attachments", []):
            extracted = att.get("extracted_text")
            if extracted and isinstance(extracted, dict):
                text = extracted.get("text", "")
                if text:
                    search_parts.append(text)

        search_text = " ".join(search_parts).lower()

        # Try direct entity name match (case-insensitive)
        for code, name in self.entity_names.items():
            if name.lower() in search_text:
                logger.debug(f"VEN-INV: Entity match by name: {code}")
                return code, name

        # Try VAT number patterns
        entity_by_vat = {
            "CHE": "CH1",
            "ATU": "AT1",
            "DE": "DE1",  # Default to DE1 for Germany (could be DE2)
            "FR": "FR1",
            "IT": "IT1",
            "GB": "UK1",
        }
        for vat_prefix, code in entity_by_vat.items():
            if re.search(rf"{vat_prefix}[\s\-]?\d", search_text):
                logger.debug(f"VEN-INV: Entity match by VAT: {code}")
                return code, self.entity_names.get(code, code)

        logger.warning("VEN-INV: Could not detect entity from email content")
        return "UNKNOWN", "Unknown"

    def _get_invoices_address(self, entity_code: str) -> str:
        """Get the invoices email address for an entity."""
        if entity_code == "CH1":
            return "invoices@eatplanted.com"
        elif entity_code == "UNKNOWN":
            return None
        else:
            return f"{entity_code.lower()}-invoices@eatplanted.com"

    def _is_in_cc(self, email_dict: Dict, invoices_addr: str) -> bool:
        """Check if invoices_addr is in CC recipients."""
        if not invoices_addr:
            return False

        cc_recipients = email_dict.get("cc_recipients", [])
        invoices_addr_lower = invoices_addr.lower()

        for recipient in cc_recipients:
            if isinstance(recipient, dict):
                email = recipient.get("email", "").lower()
            else:
                email = str(recipient).lower()

            if email == invoices_addr_lower:
                return True

        return False

    def _check_invoices_mailbox(
        self, email_dict: Dict, graph_client
    ) -> bool:
        """
        Check if email reached invoices@eatplanted.com inbox (CH1 only).

        Args:
            email_dict: Email dict
            graph_client: GraphAPIClient instance

        Returns:
            True if found in invoices@eatplanted.com, False otherwise
        """
        try:
            subject = email_dict.get("subject", "")
            from_email = email_dict.get("from", {}).get("email", "")

            if not subject or not from_email:
                logger.debug("VEN-INV: Can't check mailbox without subject or from_email")
                return False

            # Build a search filter for same sender + subject in last 7 days
            # Use OData filter: from email and subject match
            # Escape special characters in search string
            from_filter = f"from/emailAddress/address eq '{from_email}'"
            subject_filter = f"contains(subject, '{subject.replace(chr(39), chr(39) + chr(39))}')"
            search_filter = f"{from_filter} and {subject_filter}"

            # Query invoices@eatplanted.com inbox
            folder_id = graph_client.get_or_create_folder(
                "invoices@eatplanted.com", "inbox"
            )
            if not folder_id:
                logger.debug("VEN-INV: Could not access invoices@eatplanted.com")
                return False

            messages = graph_client.get_folder_messages(
                "invoices@eatplanted.com",
                folder_id,
                max_results=10,
                days_back=7
            )

            if messages:
                logger.info(
                    f"VEN-INV: Found {len(messages)} matching message(s) "
                    f"in invoices@eatplanted.com"
                )
                return True

            return False

        except Exception as e:
            logger.error(
                f"VEN-INV: Error checking invoices mailbox: {e}"
            )
            return False

    def _create_draft_reply(
        self,
        email_dict: Dict,
        entity_code: str,
        entity_name: str,
        invoices_addr: str,
        graph_client
    ) -> Optional[str]:
        """
        Create a draft reply to the vendor with correct invoices address.

        Args:
            email_dict: Email dict
            entity_code: Entity code (e.g., "CH1")
            entity_name: Entity legal name (e.g., "Planted Foods AG")
            invoices_addr: Correct invoices email address
            graph_client: GraphAPIClient instance

        Returns:
            Draft message ID if successful, None otherwise
        """
        try:
            # Build reply body
            from_name = email_dict.get("from", {}).get("name", "")
            if not from_name:
                from_name = "Dear Sender"
            else:
                from_name = f"Dear {from_name.split()[0]}"

            body_html = f"""<html>
<body>
<p>{from_name},</p>

<p>Thank you for sending us your invoice. To ensure correct and timely processing,
please direct all future invoices for <strong>{entity_name}</strong> to:</p>

<p><strong>{invoices_addr}</strong></p>

<p>We have forwarded your email internally.</p>

<p>Kind regards,<br/>
Accounting Team<br/>
Planted Foods</p>
</body>
</html>"""

            draft_id = graph_client.create_draft_reply(
                config.accounting_mailbox,
                email_dict.get("id", ""),
                body_html
            )

            if draft_id:
                logger.info(f"VEN-INV: Created draft reply: {draft_id}")
            else:
                logger.error("VEN-INV: Failed to create draft reply")

            return draft_id

        except Exception as e:
            logger.error(f"VEN-INV: Error creating draft reply: {e}", exc_info=True)
            return None

    def close(self):
        """Clean up resources (if any)."""
        pass
