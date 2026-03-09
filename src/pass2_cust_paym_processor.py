"""Pass 2 processor for CUST-PAYM emails.

Reads attachment data per case (e.g., LIDL Zahlungsavis PDF),
extracts document numbers, amounts, and payment date,
then generates a BC configuration package for journal import.
"""

import base64
import logging
import re
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Optional

import pdfplumber
import yaml

logger = logging.getLogger(__name__)


def _load_cases() -> dict:
    """Load enabled CUST-PAYM case definitions."""
    cases_file = Path(__file__).parent.parent / "config" / "cust_paym_cases.yaml"
    if not cases_file.exists():
        return {}
    with open(cases_file, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}
    cases = data.get("cases", {})
    return {k: v for k, v in cases.items() if v.get("enabled", True)}


class Pass2CustPaymProcessor:
    """Pass 2 processor for CUST-PAYM emails."""

    def __init__(self):
        self.cases = _load_cases()
        self._processed = 0

    def process_email(self, email_dict: dict, graph_client=None) -> Optional[dict]:
        """Process a CUST-PAYM email: parse attachment, extract payment data.

        Args:
            email_dict: Email dict (from email.to_dict()) with classification and attachments.
            graph_client: GraphAPIClient instance for downloading attachments if needed.

        Returns:
            pass2_results dict with extracted payment data, or None on failure.
        """
        # Determine case from classification
        classification = email_dict.get("classification", {})
        case_id = classification.get("cust_paym_case_id")
        if not case_id:
            # Try to match from sender
            case_id = self._match_case(email_dict)

        if not case_id or case_id not in self.cases:
            logger.warning(f"No CUST-PAYM case matched for email: {email_dict.get('subject', '')}")
            return self._result(case_id=None, error="No matching case found")

        case_cfg = self.cases[case_id]
        parser_name = case_cfg.get("parser", "")

        # Select parser
        if parser_name == "lidl_zahlungsavis":
            payment_data = self._parse_lidl_zahlungsavis(email_dict, graph_client)
        else:
            logger.warning(f"Unknown parser: {parser_name}")
            return self._result(case_id=case_id, error=f"Parser not implemented: {parser_name}")

        if not payment_data:
            return self._result(case_id=case_id, error="PDF parsing failed")

        self._processed += 1
        return self._result(
            case_id=case_id,
            payment_data=payment_data,
            case_cfg=case_cfg,
        )

    def _match_case(self, email_dict: dict) -> Optional[str]:
        """Match email to a case by sender."""
        from_data = email_dict.get("from", {})
        sender = (from_data.get("email") or "").lower() if isinstance(from_data, dict) else ""

        for case_id, case_cfg in self.cases.items():
            triggers = case_cfg.get("triggers", {})
            for addr in triggers.get("sender_email_exact", []):
                if addr.lower() == sender:
                    return case_id
        return None

    # ── Lidl Zahlungsavis Parser ──────────────────────────────────────────

    def _parse_lidl_zahlungsavis(self, email_dict: dict, graph_client=None) -> Optional[dict]:
        """Parse Lidl Zahlungsavis PDF attachment.

        Extracts:
        - Payment date (Fälligkeitstag)
        - Wire transfer number (Überweisung Nr.)
        - Total amount (Gesamt-Summe)
        - Line items: Beleg (Lidl doc no), Ihr Beleg (our doc no), Datum, Abzüge, Bruttobetrag
        """
        # Find the PDF attachment
        pdf_content = self._get_pdf_attachment(email_dict, graph_client)
        if not pdf_content:
            logger.error("No PDF attachment found for Lidl Zahlungsavis")
            return None

        try:
            return self._extract_lidl_data(pdf_content)
        except Exception as e:
            logger.error(f"Lidl PDF parsing failed: {e}", exc_info=True)
            return None

    def _get_pdf_attachment(self, email_dict: dict, graph_client=None) -> Optional[bytes]:
        """Get PDF attachment content, either from extracted_text or by downloading."""
        attachments = email_dict.get("attachments", [])
        pdf_att = None
        for att in attachments:
            name = att.get("name", "").lower()
            if name.endswith(".pdf"):
                pdf_att = att
                break

        if not pdf_att:
            return None

        # If we have contentBytes in the attachment dict (from Graph API direct fetch)
        if "contentBytes" in pdf_att:
            return base64.b64decode(pdf_att["contentBytes"])

        # If graph_client provided, download the attachment
        if graph_client and "id" in pdf_att:
            from .config import config
            mailbox = config.mailbox_address
            msg_id = email_dict.get("id", "")
            if msg_id:
                endpoint = f"/users/{mailbox}/messages/{msg_id}/attachments/{pdf_att['id']}"
                result = graph_client._make_request("GET", endpoint)
                if result and "contentBytes" in result:
                    return base64.b64decode(result["contentBytes"])

        logger.warning("Could not download PDF attachment")
        return None

    def _extract_lidl_data(self, pdf_content: bytes) -> dict:
        """Extract structured data from Lidl Zahlungsavis PDF bytes."""
        import tempfile
        import os

        # Write PDF to temp file for pdfplumber
        tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
        try:
            tmp.write(pdf_content)
            tmp.close()

            all_text = []
            with pdfplumber.open(tmp.name) as pdf:
                for page in pdf.pages:
                    text = page.extract_text() or ""
                    all_text.append(text)

            full_text = "\n".join(all_text)
            return self._parse_lidl_text(full_text)
        finally:
            os.unlink(tmp.name)

    def _parse_lidl_text(self, text: str) -> dict:
        """Parse extracted text from Lidl Zahlungsavis."""
        result = {
            "payment_date": None,
            "wire_transfer_no": None,
            "total_amount": None,
            "currency": "EUR",
            "line_items": [],
        }

        # Extract wire transfer number and payment date from header
        # Pattern: "Überweisung Nr. 80310906 ... Fälligkeitstag 23.09.2025"
        wire_match = re.search(r"berweisung\s+Nr\.\s+(\d+)", text)
        if wire_match:
            result["wire_transfer_no"] = wire_match.group(1)

        date_match = re.search(r"lligkeitstag\s+(\d{2}\.\d{2}\.\d{4})", text)
        if date_match:
            result["payment_date"] = date_match.group(1)

        # Extract total amount from last page
        # Pattern: "Gesamt-Summe 240.223,54" or Zahl-Betrag line
        total_match = re.search(r"Gesamt-Summe\s+([\d.,]+(?:-)?\s*)", text)
        if total_match:
            result["total_amount"] = self._parse_german_amount(total_match.group(1).strip())

        # If no Gesamt-Summe, try Zahl-Betrag
        if result["total_amount"] is None:
            zahl_match = re.search(r"Zahl-Betrag\s*\n.*?(\*+)([\d.,]+)\*", text)
            if zahl_match:
                result["total_amount"] = self._parse_german_amount(zahl_match.group(2))

        # Extract line items
        # Pattern: "11785241 RS-87221-25 09.09.2025 0,00 156,90-"
        # Beleg (Lidl doc) | Ihr Beleg (our doc) | Datum | Abzüge | Bruttobetrag
        line_pattern = re.compile(
            r"^(\d{5,12})\s+"          # Beleg (Lidl document number)
            r"(\S+)\s+"                # Ihr Beleg (our document number)
            r"(\d{2}\.\d{2}\.\d{4})\s+"  # Datum
            r"([\d.,]+)\s+"            # Abzüge
            r"([\d.,]+-?)\s*$",        # Bruttobetrag (may have trailing -)
            re.MULTILINE
        )

        for m in line_pattern.finditer(text):
            lidl_doc = m.group(1)
            our_doc = m.group(2)
            date_str = m.group(3)
            deductions = self._parse_german_amount(m.group(4))
            gross_amount = self._parse_german_amount(m.group(5))

            result["line_items"].append({
                "lidl_doc_no": lidl_doc,
                "our_doc_no": our_doc,
                "date": date_str,
                "deductions": deductions,
                "gross_amount": gross_amount,
            })

        # Skip Übertrag (carry-forward) lines — they're summaries not real items
        # Already handled by regex not matching "Übertrag" prefix

        logger.info(
            f"Lidl parser: {len(result['line_items'])} line items, "
            f"total={result['total_amount']}, date={result['payment_date']}"
        )
        return result

    @staticmethod
    def _parse_german_amount(s: str) -> float:
        """Parse German-format amount: '1.234,56' → 1234.56, '156,90-' → -156.90"""
        s = s.strip().replace("\u2019", "")  # Remove thin space used as thousands sep
        is_negative = s.endswith("-")
        s = s.rstrip("-")
        # German format: dots as thousands, comma as decimal
        s = s.replace(".", "").replace(",", ".")
        try:
            val = float(s)
            return -val if is_negative else val
        except ValueError:
            return 0.0

    # ── Result builder ────────────────────────────────────────────────────

    def _result(self, *, case_id, payment_data=None, case_cfg=None, error=None) -> dict:
        """Build pass2_results dict for CUST-PAYM."""
        r = {
            "pass2_timestamp": datetime.utcnow().isoformat() + "Z",
            "pass2_model": "rule-based",
            "cust_paym_case_id": case_id,
        }
        if case_cfg:
            r["planted_entity"] = case_cfg.get("planted_entity")
            r["principal_customer_no"] = case_cfg.get("principal_customer_no")
            r["principal_customer_name"] = case_cfg.get("principal_customer_name")
            r["bank_account_no"] = case_cfg.get("bank_account_no")
            r["currency"] = case_cfg.get("currency")
        if payment_data:
            r["payment_date"] = payment_data.get("payment_date")
            r["wire_transfer_no"] = payment_data.get("wire_transfer_no")
            r["total_amount"] = payment_data.get("total_amount")
            r["line_item_count"] = len(payment_data.get("line_items", []))
            r["payment_data"] = payment_data
        if error:
            r["error"] = error
        return r

    @property
    def processed_count(self) -> int:
        return self._processed

    def close(self):
        """Clean up resources."""
        pass
