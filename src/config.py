"""Configuration management for the accounting triage system."""

import os
from pathlib import Path
from typing import Any, Dict, Optional

import yaml
from dotenv import load_dotenv


# Load environment variables from .env file
load_dotenv()


class Config:
    """Configuration management."""

    def __init__(self):
        """Initialize configuration from environment and settings.yaml."""
        self.config_dir = Path(__file__).parent.parent / "config"
        self.settings_file = self.config_dir / "settings.yaml"
        self._settings = self._load_settings()

    def _load_settings(self) -> Dict[str, Any]:
        """Load settings from YAML file."""
        if self.settings_file.exists():
            with open(self.settings_file, "r") as f:
                return yaml.safe_load(f) or {}
        return {}

    @property
    def azure_client_id(self) -> str:
        """Azure Client ID."""
        return os.getenv("AZURE_CLIENT_ID", "")

    @property
    def azure_tenant_id(self) -> str:
        """Azure Tenant ID."""
        return os.getenv("AZURE_TENANT_ID", "")

    @property
    def azure_client_secret(self) -> str:
        """Azure Client Secret."""
        return os.getenv("AZURE_CLIENT_SECRET", "")

    @property
    def power_automate_flow_url(self) -> str:
        """Power Automate flow URL for SharePoint uploads."""
        return os.getenv("POWER_AUTOMATE_FLOW_URL", "")

    @property
    def local_folder_path(self) -> str:
        """Local folder path for email JSON files."""
        return os.getenv("LOCAL_FOLDER_PATH", "")

    @property
    def llm_provider(self) -> str:
        """LLM provider for classification: 'gemini' (default) or 'openai'."""
        return os.getenv("LLM_PROVIDER", "gemini")

    @property
    def openai_api_key(self) -> str:
        """OpenAI API key for classification (only needed if LLM_PROVIDER=openai)."""
        return os.getenv("OPENAI_API_KEY", "")

    @property
    def confluence_page_url(self) -> str:
        """Confluence page URL for category table."""
        return os.getenv("CONFLUENCE_PAGE_URL", "https://eatplanted.atlassian.net/wiki/spaces/FC/pages/2361458692")

    @property
    def confluence_email(self) -> str:
        """Confluence API email for authentication."""
        return os.getenv("CONFLUENCE_EMAIL", "")

    @property
    def confluence_api_token(self) -> str:
        """Confluence API token for authentication."""
        return os.getenv("CONFLUENCE_API_TOKEN", "")

    # --- Jira Configuration ---

    @property
    def jira_user_email(self) -> str:
        """Jira user email for Basic Auth."""
        return os.getenv("JIRA_USER_EMAIL", "")

    @property
    def jira_api_token(self) -> str:
        """Jira API token for Basic Auth."""
        return os.getenv("JIRA_API_TOKEN", "")

    # --- Fabric / Business Central Configuration ---

    @property
    def fabric_sql_endpoint(self) -> str:
        """Fabric SQL analytics endpoint URL."""
        return os.getenv("FABRIC_SQL_ENDPOINT", "")

    @property
    def fabric_database(self) -> str:
        """Fabric database (Lakehouse/Warehouse) name."""
        return os.getenv("FABRIC_DATABASE", "")

    @property
    def fabric_client_id(self) -> str:
        """Azure AD App Client ID for Fabric access."""
        return os.getenv("FABRIC_CLIENT_ID", "")

    @property
    def fabric_client_secret(self) -> str:
        """Azure AD App Client Secret for Fabric access."""
        return os.getenv("FABRIC_CLIENT_SECRET", "")

    @property
    def bc_table_config(self) -> Dict[str, Any]:
        """BC table configuration from settings.yaml."""
        return {
            "database": self.get("business_central.database", "Finance_LakeHouse_Mod"),
            "schema": self.get("business_central.schema", "dbo"),
            "table": self.get("business_central.table", "bc_vendor_ledger_entry"),
        }

    @property
    def bc_columns(self) -> Dict[str, str]:
        """BC column name mapping from settings.yaml."""
        return self.get("business_central.columns", {
            "entity": "entity",
            "external_document_no": "external_document_no",
            "vendor_no": "vendor_no",
            "vendor_name": "vendor_name",
            "amount": "amount",
            "remaining_amount": "remaining_amount",
            "due_date": "due_date",
            "posting_date": "posting_date",
            "open": "open",
            "document_type": "document_type",
            "document_no": "document_no",
            "currency_code": "currency_code",
        })

    @property
    def entity_names(self) -> Dict[str, str]:
        """Entity code to human-readable name mapping."""
        return self.get("business_central.entity_names", {
            "AT1": "Planted Foods Austria GmbH",
            "CH1": "Planted Foods AG",
            "DE1": "Planted Foods GmbH",
            "DE2": "Planted Foods Production GmbH",
            "FR1": "Planted Foods SAS",
            "IT1": "Planted Foods SRL",
            "UK1": "Eatplanted Ltd",
        })

    @property
    def accounting_mailbox(self) -> str:
        """Accounting mailbox email address."""
        return self.get("accounting_triage.mailbox", "accounting@eatplanted.com")

    @property
    def dry_run(self) -> bool:
        """Whether to run in dry-run mode (no mutations)."""
        return self.get("accounting_triage.dry_run", True)

    @property
    def max_emails(self) -> int:
        """Maximum number of emails to read."""
        return self.get("accounting_triage.email_reader.max_emails", 50)

    @property
    def days_back(self) -> int:
        """Number of days back to read emails from."""
        return self.get("accounting_triage.email_reader.days_back", 7)

    @property
    def attachment_formats(self) -> list:
        """Supported attachment formats."""
        return self.get("accounting_triage.attachments.supported_formats", [
            ".pdf", ".xlsx", ".xls", ".csv", ".png", ".jpg"
        ])

    @property
    def max_attachment_size_mb(self) -> int:
        """Max attachment size in MB."""
        return self.get("accounting_triage.attachments.max_size_mb", 25)

    def get(self, key: str, default: Any = None) -> Any:
        """Get a configuration value using dot notation."""
        keys = key.split(".")
        value = self._settings

        for k in keys:
            if isinstance(value, dict):
                value = value.get(k)
                if value is None:
                    return default
            else:
                return default

        return value if value is not None else default


# Global config instance
config = Config()
