"""Pass 2 processor for VEN-INV (vendor invoice) emails."""

import logging
import re
from datetime import datetime
from typing import Dict, Optional, List

from .config import config

logger = logging.getLogger(__name__)

# Confluence page with invoices routing instructions
INVOICES_CONFLUENCE_URL = (
    "https://eatplanted.atlassian.net/wiki/spaces/FAQ/pages/619741281/"
    "Incoming+invoices+from+suppliers+Yokoy+BLP"
)

# Patterns that indicate a no-reply sender (skip draft reply for these)
NO_REPLY_PATTERNS = [
    r"no[\-_.]?reply",
    r"noreply",
    r"do[\-_.]?not[\-_.]?reply",
    r"automated",
    r"mailer[\-_.]?daemon",
    r"postmaster",
]
_NO_REPLY_RE = re.compile("|".join(NO_REPLY_PATTERNS), re.IGNORECASE)

# Actions that mean "job done, archive it"
ARCHIVE_ACTIONS = {"ALREADY_IN_CC", "FOUND_IN_MAILBOX", "FORWARDED"}


class Pass2InvProcessor:
    """Processes VEN-INV emails: detects entity, checks CC/TO, forwards if needed, creates draft reply."""

    def __init__(self):
        """Initialize VEN-INV processor."""
        self.entity_names = config.entity_names  # {code: name} mapping

    def process_email(
        self, email_dict: Dict, graph_client=None, dry_run: bool = False
    ) -> Optional[Dict]:
        """
        Process a VEN-INV email:
          1. Detect entity
          2. If invoices address already in TO/CC → done, archive
          3. Search invoices mailbox for the same subject (last 3 days)
          4. If found → done, archive
          5. If not found → forward to invoices address
          6. If sender is not no-reply → create draft reply with correct address
             (CC any @eatplanted.com addresses involved)
          7. Archive

        Returns:
            Dictionary with pass2_results, or None on error
        """
        try:
            # Step 1: Detect entity from email content
            entity_code, entity_name = self._detect_entity(email_dict)

            if entity_code == "UNKNOWN":
                logger.warning(f"VEN-INV: Unknown entity in email {email_dict.get('subject')}")
                return self._result(
                    entity_code="UNKNOWN", entity_name="Unknown",
                    invoices_addr=None, action="UNKNOWN_ENTITY",
                )

            # Step 2: Get invoices address for this entity
            invoices_addr = self._get_invoices_address(entity_code)
            logger.info(f"VEN-INV: Entity={entity_code}, invoices_addr={invoices_addr}")

            # Step 3: Check if already in TO or CC → done, archive
            if self._is_in_recipients(email_dict, invoices_addr):
                logger.info(f"VEN-INV: {invoices_addr} already in TO/CC — nothing to do")
                return self._result(
                    entity_code=entity_code, entity_name=entity_name,
                    invoices_addr=invoices_addr, action="ALREADY_IN_CC",
                )

            # Step 4: Check if email already reached the invoices mailbox
            found_in_mailbox = False
            if graph_client and not dry_run:
                found_in_mailbox = self._check_invoices_mailbox(
                    email_dict, invoices_addr, graph_client
                )

            if found_in_mailbox:
                logger.info(f"VEN-INV: Email already in {invoices_addr} inbox — nothing to do")
                return self._result(
                    entity_code=entity_code, entity_name=entity_name,
                    invoices_addr=invoices_addr, action="FOUND_IN_MAILBOX",
                )

            # Collect @eatplanted.com people from the email (for forward CC + draft reply CC)
            planted_people = self._collect_planted_people(email_dict, invoices_addr)
            planted_emails = [p["email"] for p in planted_people]

            # Step 5: Forward (or draft-forward for CH1) to invoices address
            forwarded_to = None
            draft_forward_id = None
            is_ch1 = entity_code == "CH1"

            if graph_client and not dry_run:
                forward_to = [invoices_addr] + planted_emails
                comment = self._build_forward_comment(planted_people)

                if is_ch1:
                    # CH1: create draft forward for manual review (stays in Drafts)
                    draft_forward_id = graph_client.create_draft_forward(
                        config.accounting_mailbox,
                        email_dict.get("id", ""),
                        forward_to,
                        comment=comment,
                    )
                    if draft_forward_id:
                        logger.info(f"VEN-INV: Draft forward created for CH1 to {', '.join(forward_to)}")
                    else:
                        logger.error(f"VEN-INV: Failed to create draft forward for CH1")
                else:
                    # Non-CH1: forward immediately
                    success = graph_client.forward_message(
                        config.accounting_mailbox,
                        email_dict.get("id", ""),
                        forward_to,
                        comment=comment,
                    )
                    if success:
                        forwarded_to = invoices_addr
                        logger.info(f"VEN-INV: Forwarded to {', '.join(forward_to)}")
                    else:
                        logger.error(f"VEN-INV: Failed to forward to {invoices_addr}")
            else:
                logger.info(f"VEN-INV: [DRY RUN] Would {'draft-forward' if is_ch1 else 'forward'} to {invoices_addr}")

            # Step 6: Draft reply (skip if no-reply sender)
            draft_reply_id = None
            draft_created = False
            from_email = email_dict.get("from", {}).get("email", "")

            if self._is_no_reply(from_email):
                logger.info(f"VEN-INV: No-reply sender ({from_email}) — skipping draft reply")
            elif graph_client and not dry_run:
                draft_reply_id = self._create_draft_reply(
                    email_dict, entity_code, entity_name,
                    invoices_addr, planted_emails, graph_client,
                )
                draft_created = draft_reply_id is not None
            else:
                logger.info("VEN-INV: [DRY RUN] Would create draft reply")

            if is_ch1:
                action = "DRAFT_FORWARD_CREATED" if draft_forward_id else "DRAFT_FORWARD_FAILED"
            else:
                action = "FORWARDED" if forwarded_to else "FORWARD_FAILED"

            return self._result(
                entity_code=entity_code, entity_name=entity_name,
                invoices_addr=invoices_addr, action=action,
                forwarded_to=forwarded_to,
                draft_reply_id=draft_reply_id, draft_created=draft_created,
            )

        except Exception as e:
            logger.error(f"VEN-INV processing failed: {e}", exc_info=True)
            return None

    # ── helpers ──────────────────────────────────────────────────

    def _result(
        self, *, entity_code, entity_name, invoices_addr, action,
        forwarded_to=None, draft_reply_id=None, draft_created=False,
    ) -> Dict:
        return {
            "pass2_timestamp": datetime.utcnow().isoformat() + "Z",
            "pass2_model": "rule-based",
            "planted_entity": {"code": entity_code, "name": entity_name},
            "invoices_address": invoices_addr,
            "action_taken": action,
            "forwarded_to": forwarded_to,
            "draft_reply_id": draft_reply_id,
            "draft_reply_created": draft_created,
            "invoices": [],
            "urgency_level": None,
        }

    def _detect_entity(self, email_dict: Dict) -> tuple:
        """Detect entity from email content (fuzzy text search)."""
        search_parts = [
            email_dict.get("subject", ""),
            email_dict.get("body", ""),
        ]
        for att in email_dict.get("attachments", []):
            extracted = att.get("extracted_text")
            if extracted and isinstance(extracted, dict):
                text = extracted.get("text", "")
                if text:
                    search_parts.append(text)

        search_text = " ".join(search_parts).lower()

        # Try direct entity name match
        for code, name in self.entity_names.items():
            if name.lower() in search_text:
                logger.debug(f"VEN-INV: Entity match by name: {code}")
                return code, name

        # Try VAT number patterns
        entity_by_vat = {
            "CHE": "CH1",
            "ATU": "AT1",
            "DE": "DE1",
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

    def _get_invoices_address(self, entity_code: str) -> Optional[str]:
        """Get the invoices email address for an entity."""
        if entity_code == "CH1":
            return "invoices@eatplanted.com"
        elif entity_code == "UNKNOWN":
            return None
        else:
            return f"{entity_code.lower()}-invoices@eatplanted.com"

    def _is_in_recipients(self, email_dict: Dict, invoices_addr: str) -> bool:
        """Check if invoices_addr is in TO or CC recipients."""
        if not invoices_addr:
            return False

        target = invoices_addr.lower()
        for field_name in ("to_recipients", "cc_recipients"):
            for recipient in email_dict.get(field_name, []):
                addr = recipient.get("email", "").lower() if isinstance(recipient, dict) else str(recipient).lower()
                if addr == target:
                    return True
        return False

    def _is_no_reply(self, email_address: str) -> bool:
        """Return True if the sender looks like a no-reply address."""
        if not email_address:
            return True
        return bool(_NO_REPLY_RE.search(email_address))

    def _build_forward_comment(self, planted_people: List[Dict]) -> str:
        """Build the forward comment tagging planted people with a link to the Confluence page.

        Example: "@Stefan Haller: please ensure that invoices go to the correct address:
                  https://eatplanted.atlassian.net/wiki/..."
        """
        if not planted_people:
            return ""

        # Build "@Name1, @Name2" — use display name if available, else email local part
        names = []
        for p in planted_people:
            name = p.get("name", "").strip()
            if not name:
                name = p["email"].split("@")[0].replace(".", " ").title()
            names.append(f"@{name}")

        tag_str = ", ".join(names)
        return (
            f"{tag_str}: please ensure that invoices go to the correct address: "
            f"{INVOICES_CONFLUENCE_URL}"
        )

    def _collect_planted_people(self, email_dict: Dict, invoices_addr: str) -> List[Dict]:
        """Collect unique @eatplanted.com people from TO/CC/FROM (excluding invoices and accounting).

        Returns list of {"email": "...", "name": "..."} dicts.
        """
        excluded = {
            invoices_addr.lower() if invoices_addr else "",
            config.accounting_mailbox.lower(),
        }
        seen = set()
        people = []

        def _add(addr: str, name: str):
            addr_l = addr.lower()
            if addr_l.endswith("@eatplanted.com") and addr_l not in excluded and addr_l not in seen:
                seen.add(addr_l)
                people.append({"email": addr_l, "name": name or ""})

        # From sender
        _add(
            email_dict.get("from", {}).get("email", ""),
            email_dict.get("from", {}).get("name", ""),
        )

        # From TO and CC
        for field_name in ("to_recipients", "cc_recipients"):
            for r in email_dict.get(field_name, []):
                if isinstance(r, dict):
                    _add(r.get("email", ""), r.get("name", ""))

        return people

    def _check_invoices_mailbox(
        self, email_dict: Dict, invoices_addr: str, graph_client
    ) -> bool:
        """Check if the same email (by sender + subject) already reached the invoices mailbox."""
        try:
            subject = email_dict.get("subject", "")
            from_email = email_dict.get("from", {}).get("email", "")

            if not subject or not from_email:
                return False

            messages = graph_client.get_mailbox_messages(
                invoices_addr,
                max_results=20,
                days_back=3,
            )

            if not messages:
                return False

            # Match by sender + subject substring
            subject_lower = subject.lower().strip()
            from_lower = from_email.lower()

            for msg in messages:
                msg_from = msg.get("from", {}).get("emailAddress", {}).get("address", "").lower()
                msg_subject = msg.get("subject", "").lower().strip()
                if msg_from == from_lower and (subject_lower in msg_subject or msg_subject in subject_lower):
                    logger.info(f"VEN-INV: Found matching message in {invoices_addr}")
                    return True

            return False

        except Exception as e:
            logger.error(f"VEN-INV: Error checking invoices mailbox {invoices_addr}: {e}")
            return False

    def _create_draft_reply(
        self,
        email_dict: Dict,
        entity_code: str,
        entity_name: str,
        invoices_addr: str,
        cc_addresses: List[str],
        graph_client,
    ) -> Optional[str]:
        """Create a draft reply with the correct invoices address, CC-ing planted colleagues."""
        try:
            from_name = email_dict.get("from", {}).get("name", "")
            greeting = f"Dear {from_name.split()[0]}" if from_name else "Dear Sender"

            body_html = f"""<html>
<body>
<p>{greeting},</p>

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
                body_html,
                cc_addresses=cc_addresses,
            )

            if draft_id:
                logger.info(f"VEN-INV: Created draft reply: {draft_id}" +
                            (f" (CC: {', '.join(cc_addresses)})" if cc_addresses else ""))
            else:
                logger.error("VEN-INV: Failed to create draft reply")

            return draft_id

        except Exception as e:
            logger.error(f"VEN-INV: Error creating draft reply: {e}", exc_info=True)
            return None

    def close(self):
        """Clean up resources (if any)."""
        pass
