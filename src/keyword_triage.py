"""Keyword-based email triage (Pass 0) — no LLM calls."""

import logging
import math
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import yaml



logger = logging.getLogger(__name__)


class KeywordTriage:
    """Classifies emails using keyword/pattern matching. Zero API calls."""

    # Category ID → display name
    CATEGORY_NAMES = {
        "VEN-INV": "Vendor Invoice",
        "VEN-REM": "Vendor Payment Reminder",
        "VEN-FOLLOWUP": "Vendor Follow-up Query",
        "CUST-REM-FOLLOWUP": "Customer Follow-up to Issued Payment Reminder",
        "CUST-REMIT": "Customer Remittance Advice",
        "NO_ACTION_NEEDED": "No Action Needed",
        "OTHER": "Uncategorized / Other",
    }

    def __init__(self):
        """Load keyword rules from config/keyword_rules.yaml."""
        self.rules_file = Path(__file__).parent.parent / "config" / "keyword_rules.yaml"
        self.rules = self._load_rules()
        self.settings = self.rules.get("settings", {})
        self.categories = self.rules.get("categories", {})
        self.priority_rules = self.rules.get("priority", {})

        # Weights
        self.subject_weight = self.settings.get("subject_weight", 3.0)
        self.body_weight = self.settings.get("body_weight", 1.0)
        self.attachment_weight = self.settings.get("attachment_weight", 1.5)
        self.high_threshold = self.settings.get("high_confidence_threshold", 0.80)
        self.medium_threshold = self.settings.get("medium_confidence_threshold", 0.50)
        self.internal_domains = [
            d.lower() for d in self.settings.get("internal_domains", [])
        ]

    def _load_rules(self) -> dict:
        """Load keyword rules from YAML file."""
        if not self.rules_file.exists():
            logger.warning(f"Keyword rules not found: {self.rules_file}")
            return {}
        with open(self.rules_file, "r", encoding="utf-8") as f:
            return yaml.safe_load(f) or {}

    def classify(self, email: Dict) -> Dict:
        """
        Classify an email using keyword matching.

        Args:
            email: Email dict (from email.to_dict()) with subject, body, etc.

        Returns:
            Classification dict in the same schema as EmailClassifier,
            plus classification_method, keyword_confidence, keyword_scores.
        """
        # Normalize inputs
        subject_lower = (email.get("subject") or "").lower()
        subject_original = (email.get("subject") or "")
        body_preview_lower = (email.get("body_preview") or "").lower()
        body_lower = (email.get("body") or "").lower()
        searchable_body = body_preview_lower + " " + body_lower

        # Sender info
        from_data = email.get("from", {})
        if isinstance(from_data, dict):
            sender_email = (from_data.get("email") or "").lower()
            sender_name = (from_data.get("name") or "").lower()
        else:
            sender_email = (email.get("from_email") or "").lower()
            sender_name = (email.get("from_name") or "").lower()

        # Check rigid rules first (high-certainty matches) before probabilistic scoring
        rigid_result = self._check_rigid_rules(sender_email, subject_original)
        if rigid_result:
            return rigid_result

        # Detect if this is a reply/forward chain (subject line indicator)
        is_reply_forward = any(
            subject_lower.startswith(prefix)
            for prefix in ["re:", "fw:", "fwd:", "aw:"]
        )

        # Attachment extensions
        att_extensions = []
        for att in email.get("attachments", []):
            name = att.get("name", "")
            if "." in name:
                att_extensions.append("." + name.rsplit(".", 1)[-1].lower())

        # Score each category
        scores: Dict[str, Tuple[float, float, List[str]]] = {}
        for cat_id, cat_config in self.categories.items():
            raw_score, matched = self._score_category(
                cat_id, cat_config, subject_lower, searchable_body,
                sender_email, att_extensions, is_reply_forward,
            )
            confidence = self._score_to_confidence(raw_score)
            scores[cat_id] = (raw_score, confidence, matched)

        # Pick the winner
        if not scores:
            return self._build_result("OTHER", 0.0, [], scores, sender_email, sender_name, subject_lower, searchable_body)

        best_cat = max(scores, key=lambda k: scores[k][1])
        best_raw, best_conf, best_matched = scores[best_cat]

        # Below medium threshold → OTHER
        if best_conf < self.medium_threshold:
            return self._build_result("OTHER", best_conf, [], scores, sender_email, sender_name, subject_lower, searchable_body)

        return self._build_result(best_cat, best_conf, best_matched, scores, sender_email, sender_name, subject_lower, searchable_body)

    def _check_rigid_rules(self, sender_email: str, subject_original: str) -> Optional[Dict]:
        """
        Check rigid high-certainty rules before probabilistic scoring.
        Returns a classification dict if a rule matches, None otherwise.
        """
        rigid_rules = self.rules.get("rigid_rules", [])
        if not rigid_rules:
            return None

        subject_lower = subject_original.lower()

        for rule in rigid_rules:
            # All conditions in a rule must match (AND logic)
            match = True

            # Check sender_email_exact (exact string match)
            if "sender_email_exact" in rule:
                if rule["sender_email_exact"].lower() != sender_email:
                    match = False

            # Check sender_domain (substring match)
            if match and "sender_domain" in rule:
                if rule["sender_domain"].lower() not in sender_email:
                    match = False

            # Check subject_starts_with (any prefix match)
            if match and "subject_starts_with" in rule:
                prefixes = rule["subject_starts_with"]
                prefix_match = False
                for prefix in prefixes:
                    if subject_original.startswith(prefix):
                        prefix_match = True
                        break
                if not prefix_match:
                    match = False

            # If all conditions matched, return NO_ACTION_NEEDED with HIGH confidence
            if match:
                logger.info(f"Rigid rule matched: {rule.get('description', 'Unknown rule')}")
                return {
                    "confidence_level": "HIGH",
                    "primary_category": {
                        "id": "NO_ACTION_NEEDED",
                        "name": self.CATEGORY_NAMES.get("NO_ACTION_NEEDED", "No Action Needed"),
                    },
                    "secondary_category": None,
                    "priority": "PRIO_LOW",
                    "requires_manual_review": False,
                    "extracted_entities": {},
                    "summary": None,
                    "reasoning": f"Rigid rule matched: {rule.get('description', 'No action needed email')}",
                    "classification_method": "keyword",
                    "keyword_confidence": 1.0,
                    "keyword_scores": {},
                }

        return None

    def _score_category(
        self,
        cat_id: str,
        cat_config: dict,
        subject_lower: str,
        searchable_body: str,
        sender_email: str,
        att_extensions: List[str],
        is_reply_forward: bool = False,
    ) -> Tuple[float, List[str]]:
        """Score a single category. Returns (raw_score, matched_patterns)."""
        score = 0.0
        matched = []

        # VEN-INV requires fresh email (not reply/forward) AND must have attachments
        if cat_id == "VEN-INV":
            if is_reply_forward:
                return (-1.0, ["EXCLUDED: Reply/Forward chain (RE:/FW:/AW:)"])
            if not att_extensions:
                return (-1.0, ["EXCLUDED: No attachments (VEN-INV requires attachment)"])

        keywords = cat_config.get("keywords", {})
        exclusions = cat_config.get("exclusions", {})

        # Check subject exclusions first — hard disqualifier
        for pattern in exclusions.get("subject", []):
            if pattern.lower() in subject_lower:
                return (-1.0, [f"EXCLUDED by subject: '{pattern}'"])

        # Subject keyword matches
        for pattern in keywords.get("subject", []):
            if pattern.lower() in subject_lower:
                score += self.subject_weight
                matched.append(f"subject: '{pattern}'")

        # Body keyword matches
        for pattern in keywords.get("body", []):
            if pattern.lower() in searchable_body:
                score += self.body_weight
                matched.append(f"body: '{pattern}'")

        # Body exclusion penalties
        for pattern in exclusions.get("body", []):
            if pattern.lower() in searchable_body:
                score -= 2.0
                matched.append(f"body_exclusion: '{pattern}'")

        # Sender domain penalties
        for domain in exclusions.get("sender_domain", []):
            if domain.lower() in sender_email:
                score -= 1.0
                matched.append(f"sender_penalty: '{domain}'")

        # Attachment signals
        for ext in keywords.get("attachments", []):
            if ext in att_extensions:
                score += self.attachment_weight
                matched.append(f"attachment: '{ext}'")

        return (score, matched)

    @staticmethod
    def _score_to_confidence(score: float) -> float:
        """Convert raw score to 0.0-1.0 confidence via exponential curve.

        score=3 → ~0.50, score=6 → ~0.75, score=9 → ~0.87
        """
        if score <= 0:
            return 0.0
        return 1.0 - math.exp(-0.23 * score)

    def _assess_priority(
        self,
        sender_email: str,
        sender_name: str,
        subject_lower: str,
        searchable_body: str,
    ) -> Tuple[str, List[str]]:
        """Assess priority from keyword rules. Returns (priority_level, reasons)."""
        reasons = []
        searchable_all = subject_lower + " " + searchable_body

        # Check PRIO_HIGHEST triggers
        highest_rules = self.priority_rules.get("PRIO_HIGHEST", {})
        for group_name, patterns in highest_rules.items():
            if not isinstance(patterns, list):
                continue
            for pattern in patterns:
                if pattern.lower() in searchable_all:
                    reasons.append(f"HIGHEST: {group_name}: '{pattern}'")
                    return ("PRIO_HIGHEST", reasons)

        # Check PRIO_HIGH triggers
        high_rules = self.priority_rules.get("PRIO_HIGH", {})

        # Executive senders (check name)
        exec_senders = high_rules.get("executive_senders", [])
        for name in exec_senders:
            if name.lower() in sender_name or name.lower() in sender_email:
                reasons.append(f"HIGH: executive sender: '{name}'")
                return ("PRIO_HIGH", reasons)

        # Other PRIO_HIGH keyword groups
        for group_name, patterns in high_rules.items():
            if group_name == "executive_senders" or not isinstance(patterns, list):
                continue
            for pattern in patterns:
                if pattern.lower() in searchable_all:
                    reasons.append(f"HIGH: {group_name}: '{pattern}'")
                    return ("PRIO_HIGH", reasons)

        return ("PRIO_MEDIUM", [])

    def _build_result(
        self,
        category_id: str,
        confidence: float,
        matched_patterns: List[str],
        all_scores: Dict,
        sender_email: str,
        sender_name: str,
        subject_lower: str,
        searchable_body: str,
    ) -> Dict:
        """Build classification result matching EmailClassifier JSON schema."""
        # Confidence level
        if confidence >= self.high_threshold:
            confidence_level = "HIGH"
        elif confidence >= self.medium_threshold:
            confidence_level = "MEDIUM"
        else:
            confidence_level = "LOW"

        # Secondary category (second-highest scoring)
        sorted_cats = sorted(
            all_scores.items(), key=lambda x: x[1][1], reverse=True
        )
        secondary = None
        if len(sorted_cats) >= 2:
            sec_id = sorted_cats[1][0]
            sec_conf = sorted_cats[1][1][1]
            if sec_conf >= self.medium_threshold and sec_id != category_id:
                secondary = {
                    "id": sec_id,
                    "name": self.CATEGORY_NAMES.get(sec_id, sec_id),
                }

        # Priority assessment
        priority, priority_reasons = self._assess_priority(
            sender_email, sender_name, subject_lower, searchable_body
        )

        # Build reasoning string
        if matched_patterns:
            reasoning = (
                f"Keyword triage matched {len(matched_patterns)} pattern(s): "
                + ", ".join(matched_patterns[:5])
                + (f" (+{len(matched_patterns) - 5} more)" if len(matched_patterns) > 5 else "")
            )
        else:
            reasoning = "No keyword patterns matched."

        if priority_reasons:
            reasoning += " | Priority: " + "; ".join(priority_reasons)

        return {
            "confidence_level": confidence_level,
            "primary_category": {
                "id": category_id,
                "name": self.CATEGORY_NAMES.get(category_id, category_id),
            },
            "secondary_category": secondary,
            "priority": priority,
            "requires_manual_review": confidence_level == "LOW",
            "extracted_entities": {},
            "summary": None,
            "reasoning": reasoning,
            # Extra keyword-specific fields
            "classification_method": "keyword",
            "keyword_confidence": round(confidence, 3),
            "keyword_scores": {
                cat_id: round(vals[1], 3)
                for cat_id, vals in all_scores.items()
            },
        }
