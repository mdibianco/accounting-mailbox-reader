"""Email reading functionality."""

import logging
from dataclasses import dataclass, field, asdict
from typing import Optional, List
from datetime import datetime

from .config import config
from .graph_client import GraphAPIClient
from .attachment_analyzer import AttachmentAnalyzer, ExtractedText

logger = logging.getLogger(__name__)


@dataclass
class Attachment:
    """Email attachment metadata and content."""

    id: str
    name: str
    content_type: str
    size: int
    extracted_text: Optional[ExtractedText] = None

    def to_dict(self):
        """Convert to dictionary for JSON serialization."""
        data = {
            "id": self.id,
            "name": self.name,
            "content_type": self.content_type,
            "size": self.size,
        }
        if self.extracted_text:
            et = asdict(self.extracted_text)
            # Truncate extracted text to prevent JSON explosion from large PDFs
            if et.get("text") and len(et["text"]) > MAX_ATTACHMENT_TEXT_CHARS:
                et["text"] = et["text"][:MAX_ATTACHMENT_TEXT_CHARS] + f"\n\n[...truncated at {MAX_ATTACHMENT_TEXT_CHARS:,} chars]"
            data["extracted_text"] = et
        return data


# Size limits to prevent JSON explosion (e.g. giant PDFs, long email chains)
MAX_BODY_CHARS = 10_000          # ~10 KB — anything longer is quoted reply chains
MAX_ATTACHMENT_TEXT_CHARS = 15_000  # ~15 KB per attachment — enough for ~5-page invoice
MAX_BODY_ENGLISH_CHARS = 10_000


@dataclass
class Email:
    """Parsed email message."""

    id: str
    from_email: str
    from_name: Optional[str]
    subject: str
    received_datetime: str
    body_preview: str
    body: Optional[str] = None
    has_attachments: bool = False
    is_read: bool = False
    importance: str = "normal"
    attachments: List[Attachment] = field(default_factory=list)
    classification: Optional[dict] = None
    pass2_results: Optional[dict] = None
    body_english: Optional[str] = None
    processing_status: str = "OPEN"
    # Recipients
    to_recipients: List[dict] = field(default_factory=list)
    cc_recipients: List[dict] = field(default_factory=list)
    # Conversation threading
    web_link: Optional[str] = None  # From Graph API webLink
    graph_conversation_id: Optional[str] = None  # From Graph API conversationId
    conversation_id: Optional[str] = None
    conversation_position: int = 1
    is_latest_in_conversation: bool = True
    related_emails: List[dict] = field(default_factory=list)
    is_chain: bool = False  # True if this email is part of a multi-email conversation
    jira_issue_key: Optional[str] = None  # Jira ticket key (e.g. "FH20-123")

    @staticmethod
    def _truncate(text: Optional[str], limit: int) -> Optional[str]:
        """Truncate text to limit, appending marker if truncated."""
        if text is None or len(text) <= limit:
            return text
        return text[:limit] + f"\n\n[...truncated at {limit:,} chars]"

    def to_dict(self):
        """Convert to dictionary for JSON serialization."""
        outlook_link = self.web_link or ""
        data = {
            "id": self.id,
            "outlook_link": outlook_link,
            "processing_status": self.processing_status,
            "from": {
                "email": self.from_email,
                "name": self.from_name,
            },
            "subject": self.subject,
            "received_datetime": self.received_datetime,
            "body_preview": self.body_preview,
            "body": self._truncate(self.body, MAX_BODY_CHARS),
            "has_attachments": self.has_attachments,
            "is_read": self.is_read,
            "importance": self.importance,
            "attachments": [a.to_dict() for a in self.attachments],
        }
        if self.to_recipients:
            data["to_recipients"] = self.to_recipients
        if self.cc_recipients:
            data["cc_recipients"] = self.cc_recipients
        if self.jira_issue_key:
            data["jira_issue_key"] = self.jira_issue_key
        if self.classification:
            data["classification"] = self.classification
        if self.pass2_results:
            data["pass2_results"] = self.pass2_results
        if self.body_english:
            data["body_english"] = self._truncate(self.body_english, MAX_BODY_ENGLISH_CHARS)
        # Conversation threading
        if self.graph_conversation_id:
            data["graph_conversation_id"] = self.graph_conversation_id
        if self.conversation_id:
            data["conversation_id"] = self.conversation_id
            data["conversation_position"] = self.conversation_position
            data["is_latest_in_conversation"] = self.is_latest_in_conversation
            data["is_chain"] = self.is_chain
            if self.related_emails:
                data["related_emails"] = self.related_emails
        return data


class EmailReader:
    """Reads emails from a shared mailbox."""

    def __init__(self):
        """Initialize email reader."""
        self.graph_client = GraphAPIClient()
        self.attachment_analyzer = AttachmentAnalyzer()
        self.mailbox = config.accounting_mailbox

    def read_emails(
        self,
        max_results: Optional[int] = None,
        days_back: Optional[int] = None,
        search_query: Optional[str] = None,
        extract_attachments: bool = True,
    ) -> List[Email]:
        """
        Read emails from the accounting mailbox.

        Args:
            max_results: Maximum number of emails to read
            days_back: How many days back to read
            search_query: Optional OData search query
            extract_attachments: Whether to extract attachment content

        Returns:
            List of Email objects
        """
        if max_results is None:
            max_results = config.max_emails
        if days_back is None:
            days_back = config.days_back

        logger.info(
            f"Reading {max_results} emails from {self.mailbox} "
            f"from the last {days_back} days"
        )

        messages = self.graph_client.get_mailbox_messages(
            self.mailbox,
            max_results=max_results,
            days_back=days_back,
            search_query=search_query,
        )

        if messages is None:
            raise RuntimeError("Failed to fetch emails from mailbox (API error). Check logs for details.")

        if not messages:
            logger.info("No messages found")
            return []

        emails = []
        for msg in messages:
            try:
                email = self._parse_message(msg)

                # Get full body if needed
                if email.body is None:
                    body = self.graph_client.get_message_body(self.mailbox, email.id)
                    if body:
                        email.body = self._extract_text_from_html(body)

                # Get attachments if requested
                if email.has_attachments and extract_attachments:
                    attachments = self.graph_client.get_message_attachments(
                        self.mailbox, email.id
                    )
                    if attachments:
                        email.attachments = self._process_attachments(
                            attachments, email.id
                        )

                emails.append(email)
            except Exception as e:
                logger.error(f"Error processing message {msg.get('id')}: {e}")

        logger.info(f"Successfully read {len(emails)} emails")
        return emails

    def _parse_message(self, msg: dict) -> Email:
        """Parse a Graph API message object into an Email object."""
        from_data = msg.get("from", {}).get("emailAddress", {})

        # Parse recipients (format: [{"emailAddress": {"address": "...", "name": "..."}}])
        def parse_recipients(recipients_list):
            """Parse recipient array from Graph API."""
            result = []
            if recipients_list:
                for recipient in recipients_list:
                    email_address = recipient.get("emailAddress", {})
                    result.append({
                        "email": email_address.get("address", ""),
                        "name": email_address.get("name", "")
                    })
            return result

        to_recipients = parse_recipients(msg.get("toRecipients", []))
        cc_recipients = parse_recipients(msg.get("ccRecipients", []))

        return Email(
            id=msg.get("id", ""),
            from_email=from_data.get("address", ""),
            from_name=from_data.get("name", ""),
            subject=msg.get("subject", ""),
            received_datetime=msg.get("receivedDateTime", ""),
            body_preview=msg.get("bodyPreview", ""),
            has_attachments=msg.get("hasAttachments", False),
            is_read=msg.get("isRead", False),
            importance=msg.get("importance", "normal").lower(),
            to_recipients=to_recipients,
            cc_recipients=cc_recipients,
            web_link=msg.get("webLink"),
            graph_conversation_id=msg.get("conversationId"),
        )

    def _extract_text_from_html(self, html: str) -> str:
        """Extract plain text from HTML."""
        # Simple HTML tag removal (can be improved with BeautifulSoup if needed)
        import re

        # Remove script and style elements
        html = re.sub(r"<script.*?</script>", "", html, flags=re.DOTALL)
        html = re.sub(r"<style.*?</style>", "", html, flags=re.DOTALL)

        # Remove tags
        text = re.sub(r"<[^>]+>", "", html)

        # Decode HTML entities
        import html as html_module

        text = html_module.unescape(text)

        # Remove extra whitespace
        text = re.sub(r"\s+", " ", text)

        return text.strip()

    def _process_attachments(
        self, attachments: list, message_id: str
    ) -> List[Attachment]:
        """Download and extract text from attachments."""
        processed = []

        for att in attachments:
            att_id = att.get("id", "")
            att_name = att.get("name", "")
            att_type = att.get("contentType", "")
            att_size = att.get("size", 0)

            # Check if format is supported and size is acceptable
            from pathlib import Path

            file_ext = Path(att_name).suffix.lower()
            if file_ext not in config.attachment_formats:
                logger.debug(f"Skipping unsupported attachment: {att_name}")
                continue

            if att_size > config.max_attachment_size_mb * 1024 * 1024:
                logger.warning(
                    f"Skipping attachment {att_name} (too large: {att_size} bytes)"
                )
                continue

            # Download attachment content
            logger.debug(f"Downloading attachment: {att_name}")
            content = self.graph_client.get_attachment_content(
                self.mailbox, message_id, att_id
            )

            if content:
                # Extract text
                extracted = self.attachment_analyzer.analyze(att_name, content, att_type)
                attachment = Attachment(
                    id=att_id,
                    name=att_name,
                    content_type=att_type,
                    size=att_size,
                    extracted_text=extracted,
                )
                processed.append(attachment)
                logger.debug(
                    f"Extracted text from {att_name} "
                    f"({extracted.extraction_method})"
                )

        return processed
