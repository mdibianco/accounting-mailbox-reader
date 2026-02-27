"""Email classification using LLM (Gemini or OpenAI)."""

import json
import logging
from pathlib import Path
from typing import Dict, Optional

from .config import config

logger = logging.getLogger(__name__)


class EmailClassifier:
    """Classifies emails using Gemini (default) or OpenAI."""

    def __init__(self):
        """Initialize classifier with configured LLM provider."""
        self.provider = config.llm_provider  # "gemini" or "openai"
        self.temperature = 0.1

        # Load base prompt
        self.base_prompt = self._load_base_prompt()

        # Load categories from cache
        self.categories = self._load_categories()

        if not self.categories:
            logger.warning("No categories loaded. Run: python main.py sync-categories")

        # Initialize provider
        if self.provider == "openai":
            self._init_openai()
        else:
            self._init_gemini()

    def _init_gemini(self):
        """Initialize Gemini via API key (preferred) or CLI OAuth."""
        import os
        api_key = os.environ.get("GEMINI_API_KEY", "")
        self.model = "gemini-2.5-flash"

        if api_key:
            logger.info(f"Using Gemini ({self.model}) via API key")
        else:
            from .gemini_cli_auth import get_access_token
            get_access_token()
            logger.info(f"Using Gemini ({self.model}) via CLI OAuth")

    def _init_openai(self):
        """Initialize OpenAI client."""
        from openai import OpenAI
        api_key = config.openai_api_key
        if not api_key:
            raise ValueError("OPENAI_API_KEY not configured in environment")
        self.client = OpenAI(api_key=api_key)
        self.model = "gpt-4o-mini"
        logger.info(f"Using OpenAI ({self.model})")

    def _load_base_prompt(self) -> str:
        """Load base classification prompt from file."""
        prompt_file = Path(__file__).parent.parent / "config" / "classification_prompt.txt"
        if not prompt_file.exists():
            raise FileNotFoundError(f"Classification prompt not found: {prompt_file}")

        with open(prompt_file, "r", encoding="utf-8") as f:
            return f.read()

    def _load_categories(self) -> list:
        """Load categories from cache (synced from Confluence)."""
        cache_file = Path(__file__).parent.parent / "data" / "categories_cache.json"
        if not cache_file.exists():
            logger.warning(f"Categories cache not found: {cache_file}")
            return []

        try:
            with open(cache_file, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data.get("categories", [])
        except Exception as e:
            logger.error(f"Failed to load categories cache: {e}")
            return []

    def classify(self, email: Dict, keyword_classification: Optional[Dict] = None) -> Dict:
        """
        Classify an email using the configured LLM provider.

        Args:
            email: Email dict with subject, body, sender, etc.
            keyword_classification: Optional keyword triage result for context.

        Returns:
            Classification result dict with 'model_used' key.
        """
        try:
            user_prompt = self._build_user_prompt(email, keyword_classification)
            logger.debug(f"Classifying email: {email.get('subject', 'No subject')}")

            if self.provider == "openai":
                raw_response = self._call_openai(user_prompt)
                model_used = self.model
            else:
                raw_response, model_used = self._call_gemini(user_prompt)

            # Parse JSON response
            classification = json.loads(raw_response)
            classification["model_used"] = model_used

            logger.info(
                f"Classified as {classification.get('primary_category', {}).get('id', '?')} "
                f"| {classification.get('priority', '?')} "
                f"| confidence: {classification.get('confidence_level', '?')} "
                f"| model: {model_used}"
            )

            return classification

        except json.JSONDecodeError as e:
            logger.error(f"Failed to parse LLM response as JSON: {e}")
            logger.debug(f"Raw response: {raw_response[:500] if 'raw_response' in dir() else 'N/A'}")
            return self._get_error_classification("JSON parse error")
        except Exception as e:
            logger.error(f"Classification failed: {e}", exc_info=True)
            return self._get_error_classification(str(e))

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
                {"role": "user", "content": user_prompt}
            ],
            response_format={"type": "json_object"}
        )
        return response.choices[0].message.content

    def _build_user_prompt(self, email: Dict, keyword_classification: Optional[Dict] = None) -> str:
        """Build classification prompt with email details and attachment content."""
        # Build attachment metadata + extracted text
        attachments_str = "None"
        attachment_content_str = ""
        if email.get('has_attachments') and email.get('attachments'):
            att_meta = []
            for att in email['attachments']:
                att_meta.append(
                    f"{att.get('name', 'unknown')} ({att.get('content_type', 'unknown')}, {att.get('size', 0)} bytes)"
                )
                # Include full extracted text from attachment
                extracted = att.get('extracted_text')
                if extracted and extracted.get('success') and extracted.get('text'):
                    attachment_content_str += (
                        f"\n--- Content of {att.get('name', 'unknown')} ---\n"
                        f"{extracted['text']}\n"
                        f"--- End of {att.get('name', 'unknown')} ---\n"
                    )
            attachments_str = ", ".join(att_meta)

        categories_str = json.dumps(self.categories, indent=2)

        # Handle from field (can be dict or flat)
        from_data = email.get('from', {})
        if isinstance(from_data, dict):
            from_name = from_data.get('name', 'Unknown')
            from_email = from_data.get('email', 'unknown@unknown.com')
        else:
            from_name = email.get('from_name', 'Unknown')
            from_email = email.get('from_email', 'unknown@unknown.com')

        prompt = f"""Classify this email:

**Email Details:**
Subject: {email.get('subject', 'No subject')}
From: {from_name} <{from_email}>
Received: {email.get('received_datetime', 'Unknown date')}
Has Attachments: {email.get('has_attachments', False)}
Attachments: {attachments_str}

**Body Preview:**
{email.get('body_preview', 'No preview available')}

**Full Body:**
{email.get('body', 'No body content available')}
"""

        if attachment_content_str:
            prompt += f"""
**Attachment Content (extracted text):**
{attachment_content_str}
"""

        if keyword_classification:
            kw_cat = keyword_classification.get("primary_category", {}).get("id", "UNKNOWN")
            kw_name = keyword_classification.get("primary_category", {}).get("name", "")
            kw_conf = keyword_classification.get("confidence_level", "UNKNOWN")
            kw_score = keyword_classification.get("keyword_confidence", "N/A")
            kw_reasoning = keyword_classification.get("reasoning", "N/A")
            prompt += f"""
**Keyword Triage Pre-Classification (Pass 0):**
This email was pre-classified by keyword matching:
- Category: {kw_cat} ({kw_name})
- Confidence: {kw_conf} ({kw_score})
- Matched patterns: {kw_reasoning}

Use this as a starting hypothesis. If you agree, confirm the category. If you disagree based on the full email context, override it with your own classification and explain why in your reasoning.
"""

        prompt += f"""
**Available Categories:**
{categories_str}

Return your classification as valid JSON following the specified format.
"""
        return prompt

    def _get_error_classification(self, error_msg: str) -> Dict:
        """Return a default classification when classification fails."""
        return {
            "confidence_level": "LOW",
            "primary_category": {
                "id": "OTHER",
                "name": "Uncategorized / Other"
            },
            "secondary_category": None,
            "priority": "PRIO_LOW",
            "requires_manual_review": True,
            "extracted_entities": {},
            "reasoning": f"Classification failed: {error_msg}. Manual review required."
        }
