"""Pass 2 deep analysis processor for VEN-REM emails."""

import json
import logging
from datetime import datetime
from pathlib import Path
from typing import Dict, Optional

from .config import config
from .invoice_lookup import InvoiceLookup

logger = logging.getLogger(__name__)


class Pass2Processor:
    """Orchestrates Pass 2 deep analysis for VEN-REM emails."""

    def __init__(self):
        """Initialize Pass 2 processor."""
        self.provider = config.llm_provider
        self.temperature = 0.1
        self.base_prompt = self._load_prompt()
        self.invoice_lookup = InvoiceLookup()

        # Initialize LLM provider (same pattern as EmailClassifier)
        if self.provider == "openai":
            from openai import OpenAI
            self.client = OpenAI(api_key=config.openai_api_key)
            self.model = "gpt-4o-mini"
        else:
            import os
            api_key = os.environ.get("GEMINI_API_KEY", "")
            self.model = "gemini-2.5-flash"
            if not api_key:
                from .gemini_cli_auth import get_access_token
                get_access_token()

    def _load_prompt(self) -> str:
        """Load Pass 2 prompt from file."""
        prompt_file = (
            Path(__file__).parent.parent / "config" / "pass2_ven_rem_prompt.txt"
        )
        if not prompt_file.exists():
            raise FileNotFoundError(f"Pass 2 prompt not found: {prompt_file}")
        with open(prompt_file, "r", encoding="utf-8") as f:
            return f.read()

    def process_email(self, email_dict: Dict) -> Optional[Dict]:
        """
        Run Pass 2 analysis on a single VEN-REM email.

        Args:
            email_dict: Email dict (from email.to_dict()) with classification.

        Returns:
            pass2_results dict, or None if not applicable.
        """
        # Verify this is a VEN-REM email
        classification = email_dict.get("classification", {})
        primary_cat = classification.get("primary_category", {}).get("id", "")
        if primary_cat != "VEN-REM":
            logger.debug(f"Skipping non-VEN-REM email: {primary_cat}")
            return None

        try:
            # Step 1: LLM extraction
            logger.info("Pass 2: Running LLM extraction...")
            user_prompt = self._build_user_prompt(email_dict)

            if self.provider == "openai":
                raw_response = self._call_openai(user_prompt)
                model_used = self.model
            else:
                raw_response, model_used = self._call_gemini(user_prompt)

            extraction = json.loads(raw_response)

            # Check if LLM reclassified the email
            verified_category = extraction.get("verified_category", "VEN-REM")
            classification_verified = extraction.get("classification_verified", True)

            if not classification_verified and verified_category != "VEN-REM":
                logger.warning(
                    f"Pass 2: LLM reclassified from VEN-REM to {verified_category}"
                )
                return {
                    "pass2_timestamp": datetime.utcnow().isoformat() + "Z",
                    "pass2_model": model_used,
                    "reclassified": True,
                    "reclassified_from": "VEN-REM",
                    "reclassified_to": verified_category,
                    "verification_reasoning": extraction.get("verification_reasoning", ""),
                    "urgency_level": None,
                    "urgency_reasoning": None,
                    "planted_entity": None,
                    "invoices": [],
                    "llm_raw_extraction": extraction,
                }

            logger.info(
                f"Pass 2 LLM: verified=VEN-REM, urgency={extraction.get('urgency_level')}, "
                f"entity={extraction.get('planted_entity_code')}, "
                f"invoices={len(extraction.get('invoices', []))}"
            )

            # Step 2: BC lookup for each invoice
            entity_code = extraction.get("planted_entity_code", "UNKNOWN")
            entity_name = config.entity_names.get(entity_code, "Unknown")

            invoices_with_bc = []
            for inv in extraction.get("invoices", []):
                inv_number = inv.get("invoice_number", "")
                if inv_number:
                    bc_lookup = self.invoice_lookup.lookup_invoice(
                        entity_code=entity_code,
                        invoice_number=inv_number,
                        vendor_name=inv.get("vendor_name"),
                    )
                else:
                    bc_lookup = {
                        "status": "NO_INVOICE_NUMBER",
                        "found": False,
                        "lookup_timestamp": datetime.utcnow().isoformat() + "Z",
                    }

                inv["bc_lookup"] = bc_lookup
                invoices_with_bc.append(inv)

            # Build pass2_results
            pass2_results = {
                "pass2_timestamp": datetime.utcnow().isoformat() + "Z",
                "pass2_model": model_used,
                "classification_verified": True,
                "verified_category": "VEN-REM",
                "verification_reasoning": extraction.get("verification_reasoning", ""),
                "urgency_level": extraction.get("urgency_level"),
                "urgency_reasoning": extraction.get("urgency_reasoning", ""),
                "planted_entity": {
                    "code": entity_code,
                    "name": entity_name,
                },
                "invoices": invoices_with_bc,
                "llm_raw_extraction": extraction,
            }

            return pass2_results

        except json.JSONDecodeError as e:
            logger.error(f"Pass 2: Failed to parse LLM JSON: {e}")
            return self._get_error_result(f"LLM JSON parse error: {e}")
        except Exception as e:
            logger.error(f"Pass 2 failed: {e}", exc_info=True)
            return self._get_error_result(str(e))

    def _build_user_prompt(self, email: Dict) -> str:
        """Build the user prompt for Pass 2 extraction."""
        # Build attachment metadata + extracted text (same pattern as EmailClassifier)
        attachments_str = "None"
        attachment_content_str = ""
        if email.get("has_attachments") and email.get("attachments"):
            att_meta = []
            for att in email["attachments"]:
                att_meta.append(
                    f"{att.get('name', 'unknown')} "
                    f"({att.get('content_type', 'unknown')}, "
                    f"{att.get('size', 0)} bytes)"
                )
                extracted = att.get("extracted_text")
                if extracted and extracted.get("success") and extracted.get("text"):
                    attachment_content_str += (
                        f"\n--- Content of {att.get('name', 'unknown')} ---\n"
                        f"{extracted['text']}\n"
                        f"--- End of {att.get('name', 'unknown')} ---\n"
                    )
            attachments_str = ", ".join(att_meta)

        # Handle from field (dict or flat)
        from_data = email.get("from", {})
        if isinstance(from_data, dict):
            from_name = from_data.get("name", "Unknown")
            from_email_addr = from_data.get("email", "unknown@unknown.com")
        else:
            from_name = email.get("from_name", "Unknown")
            from_email_addr = email.get("from_email", "unknown@unknown.com")

        prompt = f"""Analyze this vendor payment reminder:

**Email Details:**
Subject: {email.get('subject', 'No subject')}
From: {from_name} <{from_email_addr}>
Received: {email.get('received_datetime', 'Unknown date')}
Attachments: {attachments_str}

**Body Preview:**
{email.get('body_preview', '')}

**Full Body:**
{email.get('body', 'No body content available')}
"""

        if attachment_content_str:
            prompt += f"""
**Attachment Content (extracted text):**
{attachment_content_str}
"""

        # Add keyword triage context for verification
        keyword_class = email.get("classification", {})
        if keyword_class:
            prompt += f"""
**Keyword Triage Result (Pass 0):**
Category: {keyword_class.get('primary_category', {}).get('id', 'UNKNOWN')}
Confidence: {keyword_class.get('confidence_level', 'UNKNOWN')} ({keyword_class.get('keyword_confidence', 'N/A')})
Method: {keyword_class.get('classification_method', 'unknown')}
Reasoning: {keyword_class.get('reasoning', 'N/A')}
"""

        prompt += """
First verify the classification (Task 0), then extract urgency, invoices, and entity (Tasks 1-3).
Return valid JSON following the specified format.
"""
        return prompt

    def _call_gemini(self, user_prompt: str) -> tuple:
        """Call Gemini API with model cascade fallback.
        Returns (response_text, model_used)."""
        from .gemini_cli_auth import call_gemini_cascade
        return call_gemini_cascade(
            prompt=user_prompt,
            system_instruction=self.base_prompt,
            temperature=self.temperature,
            json_output=True,
            preferred_model=self.model,
        )

    def _call_openai(self, user_prompt: str) -> str:
        """Call OpenAI API."""
        response = self.client.chat.completions.create(
            model=self.model,
            temperature=self.temperature,
            messages=[
                {"role": "system", "content": self.base_prompt},
                {"role": "user", "content": user_prompt},
            ],
            response_format={"type": "json_object"},
        )
        return response.choices[0].message.content

    def _get_error_result(self, error_msg: str) -> Dict:
        """Return error result when Pass 2 fails."""
        return {
            "pass2_timestamp": datetime.utcnow().isoformat() + "Z",
            "pass2_model": self.model,
            "error": error_msg,
            "urgency_level": None,
            "urgency_reasoning": None,
            "planted_entity": None,
            "invoices": [],
        }

    def close(self):
        """Clean up resources."""
        if self.invoice_lookup:
            self.invoice_lookup.close()
