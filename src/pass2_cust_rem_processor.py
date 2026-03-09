"""Pass 2 processor for CUST-REM-FOLLOWUP emails.

Extracts reminder number (CREP pattern) and sender domain.
If no CREP found, uses LLM to re-verify classification.
"""

import json
import logging
import re
from datetime import datetime
from pathlib import Path
from typing import Optional

from .config import config

logger = logging.getLogger(__name__)

# Regex for planted reminder numbers (e.g. CREP01002118)
_CREP_PATTERN = re.compile(r"CREP\w+", re.IGNORECASE)


class Pass2CustRemProcessor:
    """Pass 2 processor for CUST-REM-FOLLOWUP emails."""

    def __init__(self):
        self._llm_initialized = False
        self._llm_calls = 0

    def _init_llm(self):
        """Lazy-init LLM only when needed (no CREP found)."""
        if self._llm_initialized:
            return
        self.provider = config.llm_provider
        self.temperature = 0.1
        self.system_prompt = self._build_system_prompt()

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

        self._llm_initialized = True

    def _build_system_prompt(self) -> str:
        """Load classification prompt (categories only) for reclassification."""
        prompt_file = Path(__file__).parent.parent / "config" / "classification_prompt.txt"
        if not prompt_file.exists():
            raise FileNotFoundError(f"Classification prompt not found: {prompt_file}")
        with open(prompt_file, "r", encoding="utf-8") as f:
            full_prompt = f.read()

        # Strip output format section to avoid schema conflict
        marker = "# Output Format"
        idx = full_prompt.find(marker)
        categories_only = full_prompt[:idx].rstrip() if idx != -1 else full_prompt

        return (
            categories_only
            + "\n\n---\n\n"
            + "# TASK: Re-verify classification\n\n"
            + "This email was pre-classified as CUST-REM-FOLLOWUP but lacks a planted "
            + "reminder number (CREP...) in the subject, which is unusual.\n\n"
            + "Re-examine the email and decide:\n"
            + "1. If it IS a customer responding to a planted payment reminder → keep CUST-REM-FOLLOWUP\n"
            + "2. If it is something else → reclassify to the correct category\n\n"
            + "Return JSON: {\"verified_category\": \"CATEGORY_ID\", \"classification_verified\": true/false, "
            + "\"reasoning\": \"...\"}\n"
            + "- classification_verified=true means CUST-REM-FOLLOWUP is correct\n"
            + "- classification_verified=false means reclassify to verified_category\n"
        )

    def process_email(self, email_dict: dict) -> Optional[dict]:
        """Extract reminder number and domain. LLM re-verify if no CREP found."""
        try:
            subject = email_dict.get("subject", "")
            reminder_number = self._extract_reminder_number(subject)
            sender_domain = self._extract_sender_domain(email_dict)

            if reminder_number:
                # CREP found — classification is confirmed, no LLM needed
                return self._result(
                    reminder_number=reminder_number,
                    sender_domain=sender_domain,
                    classification_verified=True,
                    model_used="rule-based",
                )

            # No CREP — ask LLM to re-verify classification
            logger.info("No CREP in subject, running LLM re-verification...")
            self._init_llm()
            self._llm_calls += 1

            verification = self._llm_verify(email_dict)
            if verification and not verification.get("classification_verified", True):
                # LLM says this is NOT CUST-REM-FOLLOWUP
                new_cat = verification.get("verified_category", "OTHER")
                logger.warning(f"LLM reclassified CUST-REM-FOLLOWUP -> {new_cat}")
                return {
                    "pass2_timestamp": datetime.utcnow().isoformat() + "Z",
                    "pass2_model": getattr(self, "model", "rule-based"),
                    "reclassified": True,
                    "reclassified_from": "CUST-REM-FOLLOWUP",
                    "reclassified_to": new_cat,
                    "verification_reasoning": verification.get("reasoning", ""),
                    "cust_reminder_number": None,
                    "cust_sender_domain": sender_domain,
                }

            # LLM confirmed CUST-REM-FOLLOWUP (just no CREP in subject)
            return self._result(
                reminder_number=None,
                sender_domain=sender_domain,
                classification_verified=True,
                model_used=getattr(self, "model", "rule-based"),
                reasoning=verification.get("reasoning", "") if verification else "",
            )

        except Exception as e:
            logger.error(f"CUST-REM-FOLLOWUP Pass 2 failed: {e}", exc_info=True)
            return self._result(
                reminder_number=None,
                sender_domain=self._extract_sender_domain(email_dict),
                classification_verified=True,
                model_used="error",
                error=str(e),
            )

    def _extract_reminder_number(self, subject: str) -> Optional[str]:
        """Extract CREP... pattern from subject."""
        match = _CREP_PATTERN.search(subject)
        return match.group(0) if match else None

    def _extract_sender_domain(self, email_dict: dict) -> Optional[str]:
        """Extract domain from sender email address."""
        from_data = email_dict.get("from", {})
        email_addr = from_data.get("email", "") if isinstance(from_data, dict) else ""
        if "@" in email_addr:
            return email_addr.split("@", 1)[1].lower()
        return None

    def _result(self, *, reminder_number, sender_domain, classification_verified,
                model_used, reasoning="", error=None) -> dict:
        """Build pass2_results dict."""
        result = {
            "pass2_timestamp": datetime.utcnow().isoformat() + "Z",
            "pass2_model": model_used,
            "classification_verified": classification_verified,
            "verified_category": "CUST-REM-FOLLOWUP",
            "cust_reminder_number": reminder_number,
            "cust_sender_domain": sender_domain,
        }
        if reasoning:
            result["verification_reasoning"] = reasoning
        if error:
            result["error"] = error
        return result

    def _llm_verify(self, email_dict: dict) -> Optional[dict]:
        """Call LLM to re-verify classification."""
        from_data = email_dict.get("from", {})
        if isinstance(from_data, dict):
            from_str = f"{from_data.get('name', '')} <{from_data.get('email', '')}>"
        else:
            from_str = str(from_data)

        user_prompt = (
            f"Re-verify this email classified as CUST-REM-FOLLOWUP:\n\n"
            f"Subject: {email_dict.get('subject', '')}\n"
            f"From: {from_str}\n"
            f"Body:\n{email_dict.get('body_preview', '')}\n\n"
            f"{email_dict.get('body', '')[:2000]}\n\n"
            f"Return JSON with verified_category, classification_verified, reasoning."
        )

        try:
            if self.provider == "openai":
                raw = self._call_openai(user_prompt)
            else:
                raw, _ = self._call_gemini(user_prompt)
            return json.loads(raw)
        except Exception as e:
            logger.error(f"LLM verification failed: {e}")
            return None

    def _call_gemini(self, user_prompt: str) -> tuple:
        """Call Gemini with cascade fallback."""
        from .gemini_cli_auth import call_gemini_cascade
        return call_gemini_cascade(
            prompt=user_prompt,
            system_instruction=self.system_prompt,
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
                {"role": "system", "content": self.system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            response_format={"type": "json_object"},
        )
        return response.choices[0].message.content

    @property
    def llm_calls(self) -> int:
        return self._llm_calls

    def close(self):
        """Clean up resources."""
        pass
