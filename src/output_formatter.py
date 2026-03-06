"""Output formatting (JSON and console)."""

import json
import logging
from typing import List, Optional
from datetime import datetime
from io import StringIO

try:
    from tabulate import tabulate
except ImportError:
    tabulate = None

from .email_reader import Email
from .graph_client import GraphAPIClient

logger = logging.getLogger(__name__)

# Hard cap: individual email JSONs should never exceed this (bytes)
MAX_JSON_SIZE_BYTES = 200_000  # 200 KB


class OutputFormatter:
    """Formats email data for output (JSON, console table, etc.)."""

    @staticmethod
    def to_json(emails: List[Email], pretty: bool = True) -> str:
        """
        Format emails as JSON.

        Args:
            emails: List of Email objects
            pretty: Whether to pretty-print JSON

        Returns:
            JSON string
        """
        data = {
            "timestamp": datetime.utcnow().isoformat(),
            "count": len(emails),
            "emails": [email.to_dict() for email in emails],
        }

        if pretty:
            return json.dumps(data, indent=2)
        else:
            return json.dumps(data)

    @staticmethod
    def to_console_table(
        emails: List[Email], include_attachments: bool = True
    ) -> str:
        """
        Format emails as a console table with key information.

        Args:
            emails: List of Email objects
            include_attachments: Include attachment information

        Returns:
            Formatted string for console output
        """
        if not tabulate:
            return OutputFormatter._to_simple_text(emails, include_attachments)

        headers = [
            "From (Domain)",
            "Subject",
            "Received",
            "Att",
            "Preview",
        ]
        rows = []

        for email in emails:
            # Extract domain from email
            from_domain = email.from_email.split("@")[1] if "@" in email.from_email else email.from_email
            
            # Count attachments
            attachment_count = len(email.attachments) if email.attachments else 0
            att_str = str(attachment_count) if attachment_count > 0 else "-"
            
            # Get preview text (prefer body preview, fallback to subject)
            preview = email.body_preview or email.subject
            preview = preview[:60] + "..." if len(preview) > 60 else preview
            
            rows.append([
                from_domain,
                email.subject[:40] + "..." if len(email.subject) > 40 else email.subject,
                email.received_datetime[:10],
                att_str,
                preview,
            ])

        output = StringIO()
        output.write(f"\nEmail Overview ({len(emails)} emails)\n")
        output.write(f"{'='*140}\n")
        output.write(tabulate(rows, headers=headers, tablefmt="grid", maxcolwidths=[15, 35, 12, 5, 65]))
        output.write(f"\n{'='*140}\n")
        
        # Summary statistics
        total_attachments = sum(len(e.attachments) if e.attachments else 0 for e in emails)
        attachment_types = {}
        for email in emails:
            if email.attachments:
                for att in email.attachments:
                    ext = att.name.split(".")[-1].lower() if "." in att.name else "unknown"
                    attachment_types[ext] = attachment_types.get(ext, 0) + 1
        
        output.write(f"Summary: {total_attachments} attachments")
        if attachment_types:
            ext_str = ", ".join([f"{count} {ext.upper()}" for ext, count in sorted(attachment_types.items(), key=lambda x: x[1], reverse=True)])
            output.write(f" ({ext_str})")
        output.write("\n")
        
        return output.getvalue()

    @staticmethod
    def _to_simple_text(emails: List[Email], include_attachments: bool = True) -> str:
        """Simple text formatting when tabulate is not available."""
        output = StringIO()

        output.write(f"\n{'='*80}\n")
        output.write(f"EMAIL REPORT - {datetime.utcnow().isoformat()}\n")
        output.write(f"{'='*80}\n")
        output.write(f"Total Emails: {len(emails)}\n\n")

        for idx, email in enumerate(emails, 1):
            output.write(f"\n{'-'*80}\n")
            output.write(f"[{idx}] {email.subject}\n")
            output.write(f"{'-'*80}\n")
            output.write(f"From:     {email.from_name} <{email.from_email}>\n")
            output.write(f"Date:     {email.received_datetime}\n")
            output.write(f"Priority: {email.importance.upper()}\n")
            output.write(f"Read:     {'Yes' if email.is_read else 'No'}\n")

            if email.body_preview:
                output.write(
                    f"Preview:  {email.body_preview[:100]}...\n"
                    if len(email.body_preview) > 100
                    else f"Preview:  {email.body_preview}\n"
                )

            if include_attachments and email.attachments:
                output.write(f"\nAttachments ({len(email.attachments)}):\n")
                for att in email.attachments:
                    output.write(
                        f"  - {att.name} ({att.content_type}, {att.size} bytes)\n"
                    )
                    if att.extracted_text:
                        if att.extracted_text.success:
                            text_preview = att.extracted_text.text[:200].replace(
                                "\n", " "
                            )
                            output.write(f"    Content: {text_preview}...\n")
                        else:
                            output.write(
                                f"    Error: {att.extracted_text.error}\n"
                            )

        output.write(f"\n{'='*80}\n")
        return output.getvalue()

    @staticmethod
    def to_detailed_text(
        emails: List[Email],
        include_body: bool = True,
        include_extracted_text: bool = True,
    ) -> str:
        """
        Format emails with full details.

        Args:
            emails: List of Email objects
            include_body: Include full email body
            include_extracted_text: Include extracted attachment text

        Returns:
            Detailed formatted string
        """
        output = StringIO()

        output.write(f"\n{'='*100}\n")
        output.write(f"DETAILED EMAIL REPORT - {datetime.utcnow().isoformat()}\n")
        output.write(f"{'='*100}\n")
        output.write(f"Total Emails: {len(emails)}\n")

        for idx, email in enumerate(emails, 1):
            output.write(f"\n{'#'*100}\n")
            output.write(f"EMAIL #{idx}\n")
            output.write(f"{'#'*100}\n")

            output.write(f"\nSubject:      {email.subject}\n")
            output.write(f"From:         {email.from_name} <{email.from_email}>\n")
            output.write(f"Date:         {email.received_datetime}\n")
            output.write(f"Priority:     {email.importance.upper()}\n")
            output.write(f"Read:         {'Yes' if email.is_read else 'No'}\n")
            output.write(f"Has Attachments: {'Yes' if email.has_attachments else 'No'}\n")

            if include_body and email.body:
                output.write(f"\n{'-'*100}\n")
                output.write("FULL BODY:\n")
                output.write(f"{'-'*100}\n")
                output.write(email.body)
                output.write("\n")

            if email.attachments:
                output.write(f"\n{'-'*100}\n")
                output.write(f"ATTACHMENTS ({len(email.attachments)}):\n")
                output.write(f"{'-'*100}\n")

                for att in email.attachments:
                    output.write(f"\n[ATTACHMENT] {att.name}\n")
                    output.write(f"  Type: {att.content_type}\n")
                    output.write(f"  Size: {att.size} bytes\n")

                    if att.extracted_text:
                        if att.extracted_text.success and include_extracted_text:
                            output.write(f"  Extraction Method: {att.extracted_text.extraction_method}\n")
                            output.write(f"  Extracted Text:\n")
                            output.write(f"  {'-'*96}\n")
                            # Indent extracted text
                            for line in att.extracted_text.text.split("\n"):
                                output.write(f"  {line}\n")
                            output.write(f"  {'-'*96}\n")
                        elif not att.extracted_text.success:
                            output.write(f"  Extraction Failed: {att.extracted_text.error}\n")

        output.write(f"\n{'='*100}\n")
        return output.getvalue()

    @staticmethod
    def save_to_file(content: str, filepath: str) -> bool:
        """
        Save formatted output to a file.

        Args:
            content: Content to save
            filepath: Path to save to

        Returns:
            True if successful
        """
        try:
            with open(filepath, "w", encoding="utf-8") as f:
                f.write(content)
            logger.info(f"Output saved to {filepath}")
            return True
        except Exception as e:
            logger.error(f"Failed to save to {filepath}: {e}")
            return False

    @staticmethod
    def save_to_local_folder(
        emails: List[Email],
        folder_path: str = r"C:\Users\MatthiasDiBianco\emails\emails",
    ) -> dict:
        """
        Save individual email JSONs to a local folder.

        Args:
            emails: List of Email objects
            folder_path: Local folder path to save JSON files

        Returns:
            Dict with upload statistics
        """
        from pathlib import Path

        stats = {
            "total": len(emails),
            "successful": 0,
            "failed": 0,
            "errors": []
        }

        # Create folder if it doesn't exist
        folder = Path(folder_path)
        try:
            folder.mkdir(parents=True, exist_ok=True)
            logger.info(f"Saving {len(emails)} emails to: {folder_path}")
        except Exception as e:
            error_msg = f"Failed to create folder {folder_path}: {e}"
            logger.error(error_msg)
            stats["errors"].append(error_msg)
            stats["failed"] = stats["total"]
            return stats

        for email in emails:
            try:
                # Create unique filename from email ID hash and timestamp
                import hashlib
                email_hash = hashlib.md5(email.id.encode()).hexdigest()[:12]
                email_timestamp = email.received_datetime.replace(":", "-").replace("T", "_").split(".")[0]
                filename = f"{email_timestamp}_{email_hash}.json"
                file_path = folder / filename

                # Convert email to JSON
                email_json = json.dumps(email.to_dict(), indent=2)

                # Safety net: warn and truncate if still over hard cap
                json_bytes = len(email_json.encode("utf-8"))
                if json_bytes > MAX_JSON_SIZE_BYTES:
                    logger.warning(
                        f"⚠ {filename} is {json_bytes:,} bytes (cap: {MAX_JSON_SIZE_BYTES:,}). "
                        f"Truncating attachment texts further."
                    )
                    email_dict = json.loads(email_json)
                    for att in email_dict.get("attachments", []):
                        et = att.get("extracted_text")
                        if et and et.get("text") and len(et["text"]) > 2000:
                            et["text"] = et["text"][:2000] + "\n\n[...hard-truncated to fit 200KB cap]"
                    if email_dict.get("body") and len(email_dict["body"]) > 5000:
                        email_dict["body"] = email_dict["body"][:5000] + "\n\n[...hard-truncated to fit 200KB cap]"
                    email_json = json.dumps(email_dict, indent=2)

                # Write to file
                with open(file_path, "w", encoding="utf-8") as f:
                    f.write(email_json)

                stats["successful"] += 1
                logger.info(f"✓ Saved: {filename} ({json_bytes:,} bytes)")

            except Exception as e:
                stats["failed"] += 1
                error_msg = f"Error saving email {email.id}: {str(e)}"
                stats["errors"].append(error_msg)
                logger.error(error_msg)

        logger.info(
            f"Local save complete: {stats['successful']}/{stats['total']} successful, "
            f"{stats['failed']} failed"
        )

        return stats

    @staticmethod
    def save_via_power_automate(
        emails: List[Email],
        flow_url: str,
    ) -> dict:
        """
        Save individual email JSONs to SharePoint via Power Automate.

        Args:
            emails: List of Email objects
            flow_url: Power Automate HTTP trigger URL

        Returns:
            Dict with upload statistics
        """
        import requests

        stats = {
            "total": len(emails),
            "successful": 0,
            "failed": 0,
            "errors": []
        }

        logger.info(f"Uploading {len(emails)} emails via Power Automate...")

        for email in emails:
            try:
                # Create unique filename from email ID hash and timestamp
                import hashlib
                email_hash = hashlib.md5(email.id.encode()).hexdigest()[:12]
                email_timestamp = email.received_datetime.replace(":", "-").replace("T", "_").split(".")[0]
                filename = f"{email_timestamp}_{email_hash}.json"

                # Prepare payload for Power Automate
                payload = {
                    "filename": filename,
                    "content": email.to_dict()
                }

                # Send to Power Automate
                response = requests.post(
                    flow_url,
                    json=payload,
                    headers={"Content-Type": "application/json"},
                    timeout=30
                )
                response.raise_for_status()

                stats["successful"] += 1
                logger.info(f"✓ Uploaded: {filename}")

            except Exception as e:
                stats["failed"] += 1
                error_msg = f"Error uploading email {email.id}: {str(e)}"
                stats["errors"].append(error_msg)
                logger.error(error_msg)

        logger.info(
            f"Power Automate upload complete: {stats['successful']}/{stats['total']} successful, "
            f"{stats['failed']} failed"
        )

        return stats

    @staticmethod
    def save_to_sharepoint(
        emails: List[Email],
        graph_client: Optional[GraphAPIClient] = None,
        site_hostname: str = "planted.sharepoint.com",
        site_path: str = "/sites/finance",
        drive_name: str = "Shared Documents",
        folder_path: str = "Accounting/Mailbox/emails",
    ) -> dict:
        """
        Save individual email JSONs to SharePoint.

        Args:
            emails: List of Email objects
            graph_client: Optional existing GraphAPIClient instance (to reuse authentication)
            site_hostname: SharePoint hostname (default: planted.sharepoint.com)
            site_path: SharePoint site path (default: /sites/finance)
            drive_name: SharePoint drive name (default: Shared Documents)
            folder_path: Folder path within drive (default: Accounting/Mailbox/emails)

        Returns:
            Dict with upload statistics: {
                "total": int,
                "successful": int,
                "failed": int,
                "errors": List[str]
            }
        """
        # Reuse existing client if provided, otherwise create new one
        if graph_client is None:
            graph_client = GraphAPIClient()

        stats = {
            "total": len(emails),
            "successful": 0,
            "failed": 0,
            "errors": []
        }

        logger.info(
            f"Uploading {len(emails)} emails to SharePoint: "
            f"{site_hostname}{site_path}/{folder_path}"
        )

        for email in emails:
            try:
                # Create unique filename from email ID hash and timestamp
                # Format: {timestamp}_{hash}.json
                import hashlib
                email_hash = hashlib.md5(email.id.encode()).hexdigest()[:12]
                email_timestamp = email.received_datetime.replace(":", "-").replace("T", "_").split(".")[0]
                filename = f"{email_timestamp}_{email_hash}.json"

                # Convert email to JSON (single email, not wrapped in array)
                email_json = json.dumps(email.to_dict(), indent=2)
                file_content = email_json.encode("utf-8")

                # Safety net: truncate if over hard cap
                if len(file_content) > MAX_JSON_SIZE_BYTES:
                    logger.warning(
                        f"⚠ {filename} is {len(file_content):,} bytes (cap: {MAX_JSON_SIZE_BYTES:,}). "
                        f"Truncating attachment texts further."
                    )
                    email_dict = json.loads(email_json)
                    for att in email_dict.get("attachments", []):
                        et = att.get("extracted_text")
                        if et and et.get("text") and len(et["text"]) > 2000:
                            et["text"] = et["text"][:2000] + "\n\n[...hard-truncated to fit 200KB cap]"
                    if email_dict.get("body") and len(email_dict["body"]) > 5000:
                        email_dict["body"] = email_dict["body"][:5000] + "\n\n[...hard-truncated to fit 200KB cap]"
                    email_json = json.dumps(email_dict, indent=2)
                    file_content = email_json.encode("utf-8")

                # Upload to SharePoint
                result = graph_client.upload_to_sharepoint(
                    site_hostname=site_hostname,
                    site_path=site_path,
                    drive_name=drive_name,
                    folder_path=folder_path,
                    file_name=filename,
                    file_content=file_content,
                )

                if result:
                    stats["successful"] += 1
                    logger.info(f"✓ Uploaded: {filename}")
                else:
                    stats["failed"] += 1
                    error_msg = f"Failed to upload {filename}"
                    stats["errors"].append(error_msg)
                    logger.error(error_msg)

            except Exception as e:
                stats["failed"] += 1
                error_msg = f"Error uploading email {email.id}: {str(e)}"
                stats["errors"].append(error_msg)
                logger.error(error_msg)

        logger.info(
            f"SharePoint upload complete: {stats['successful']}/{stats['total']} successful, "
            f"{stats['failed']} failed"
        )

        return stats
