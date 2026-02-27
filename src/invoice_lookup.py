"""Invoice lookup via Fabric Lakehouse SQL endpoint."""

import logging
import re
import struct
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

import msal

from .config import config

logger = logging.getLogger(__name__)


class InvoiceLookup:
    """Look up invoices in BC data via Fabric Lakehouse SQL."""

    def __init__(self):
        """Initialize Fabric SQL client."""
        self.sql_endpoint = config.fabric_sql_endpoint
        self.database = config.fabric_database
        self.client_id = config.fabric_client_id
        self.client_secret = config.fabric_client_secret
        self.tenant_id = config.azure_tenant_id
        self.scopes = ["https://database.windows.net/.default"]

        # Table config from settings.yaml
        self.table_config = config.bc_table_config
        self.columns = config.bc_columns

        # Full table reference
        self.full_table = (
            f"[{self.table_config['database']}]"
            f".[{self.table_config['schema']}]"
            f".[{self.table_config['table']}]"
        )

        # MSAL token cache (separate from Graph cache)
        self.cache_file = (
            Path.home() / ".accounting_mailbox_reader" / "fabric_token_cache.bin"
        )
        self.cache_file.parent.mkdir(parents=True, exist_ok=True)
        self.token_cache = msal.SerializableTokenCache()

        if self.cache_file.exists():
            with open(self.cache_file, "r") as f:
                self.token_cache.deserialize(f.read())

        self.app = None
        self._connection = None

        if self.is_configured:
            self._initialize_auth()

    @property
    def is_configured(self) -> bool:
        """Check if Fabric SQL credentials are configured."""
        return bool(
            self.sql_endpoint
            and self.database
            and self.client_id
            and self.client_secret
            and self.tenant_id
        )

    def _initialize_auth(self):
        """Initialize MSAL for Fabric (client credentials only)."""
        self.app = msal.ConfidentialClientApplication(
            client_id=self.client_id,
            client_credential=self.client_secret,
            authority=f"https://login.microsoftonline.com/{self.tenant_id}",
            token_cache=self.token_cache,
        )
        logger.info("Fabric SQL authentication initialized")

    def _save_cache(self):
        """Save token cache to disk."""
        if self.token_cache.has_state_changed:
            with open(self.cache_file, "w") as f:
                f.write(self.token_cache.serialize())

    def _get_token(self) -> Optional[str]:
        """Get access token for Fabric SQL endpoint."""
        if not self.app:
            return None

        result = self.app.acquire_token_for_client(scopes=self.scopes)

        if result and "access_token" in result:
            self._save_cache()
            return result["access_token"]
        else:
            error = result.get("error_description", result.get("error", "Unknown"))
            logger.error(f"Fabric token acquisition failed: {error}")
            return None

    def _get_connection(self):
        """Get or create pyodbc connection to Fabric SQL endpoint."""
        if self._connection:
            try:
                # Test if connection is still alive
                self._connection.execute("SELECT 1")
                return self._connection
            except Exception:
                self._connection = None

        try:
            import pyodbc
        except ImportError:
            logger.error(
                "pyodbc not installed. Run: pip install pyodbc"
            )
            return None

        token = self._get_token()
        if not token:
            return None

        # Encode token for SQL Server ODBC driver
        token_bytes = token.encode("UTF-16-LE")
        token_struct = struct.pack(
            f"<I{len(token_bytes)}s", len(token_bytes), token_bytes
        )

        conn_str = (
            f"Driver={{ODBC Driver 18 for SQL Server}};"
            f"Server={self.sql_endpoint};"
            f"Database={self.database};"
            f"Encrypt=yes;"
            f"TrustServerCertificate=no"
        )

        try:
            # SQL_COPT_SS_ACCESS_TOKEN = 1256
            self._connection = pyodbc.connect(
                conn_str, attrs_before={1256: token_struct}
            )
            logger.info("Connected to Fabric SQL endpoint")
            return self._connection
        except Exception as e:
            logger.error(f"Fabric SQL connection failed: {e}")
            return None

    def lookup_invoice(
        self,
        entity_code: str,
        invoice_number: str,
        vendor_name: Optional[str] = None,
    ) -> Dict:
        """
        Look up an invoice in BC vendor ledger entries via Fabric SQL.

        Args:
            entity_code: Planted entity code (AT1, CH1, DE1, etc.)
            invoice_number: Vendor's invoice/document number
            vendor_name: Optional vendor name for fuzzy matching

        Returns:
            bc_lookup dict with status and BC data.
        """
        timestamp = datetime.utcnow().isoformat() + "Z"

        # Check configuration
        if not self.is_configured:
            return {
                "status": "LOOKUP_NOT_CONFIGURED",
                "found": False,
                "error": "Fabric SQL credentials not configured",
                "lookup_timestamp": timestamp,
            }

        conn = self._get_connection()
        if not conn:
            return {
                "status": "LOOKUP_ERROR",
                "found": False,
                "error": "Could not connect to Fabric SQL endpoint",
                "lookup_timestamp": timestamp,
            }

        col = self.columns

        # Try exact match first
        result = self._query_exact(conn, entity_code, invoice_number, col)

        # If no result, try fuzzy match
        match_type = "exact"
        if not result:
            normalized = self._normalize_invoice_number(invoice_number)
            if normalized != invoice_number:
                result = self._query_fuzzy(
                    conn, entity_code, normalized, col
                )
                match_type = "fuzzy"

        if not result:
            return {
                "status": "NOT_FOUND",
                "found": False,
                "lookup_timestamp": timestamp,
            }

        # Determine status from remaining_amount
        remaining = result.get("remaining_amount", 0) or 0
        is_open = result.get("open", True)

        if not is_open or remaining == 0:
            status = "SETTLED"
        else:
            status = "BOOKED"

        return {
            "status": status,
            "found": True,
            "match_type": match_type,
            "external_document_no": result.get("external_document_no"),
            "document_no": result.get("document_no"),
            "vendor_number": result.get("vendor_no"),
            "vendor_name": result.get("vendor_name"),
            "posting_date": str(result.get("posting_date", "")),
            "due_date": str(result.get("due_date", "")),
            "amount": result.get("amount"),
            "remaining_amount": remaining,
            "currency_code": result.get("currency_code", ""),
            "open": is_open,
            "on_hold": result.get("on_hold", ""),
            "lookup_timestamp": timestamp,
        }

    def _query_exact(
        self, conn, entity_code: str, invoice_number: str, col: Dict
    ) -> Optional[Dict]:
        """Query for exact match on external_document_no."""
        sql = f"""
            SELECT TOP 1
                {col['external_document_no']},
                {col['document_no']},
                {col['vendor_no']},
                {col['vendor_name']},
                {col['amount']},
                {col['remaining_amount']},
                {col['due_date']},
                {col['posting_date']},
                {col['open']},
                {col['currency_code']},
                {col.get('on_hold', 'on_hold')}
            FROM {self.full_table}
            WHERE {col['entity']} = ?
              AND {col['external_document_no']} = ?
              AND {col['document_type']} = 'Invoice'
            ORDER BY {col['posting_date']} DESC
        """

        try:
            cursor = conn.cursor()
            cursor.execute(sql, (entity_code, invoice_number))
            row = cursor.fetchone()
            cursor.close()

            if row:
                return {
                    "external_document_no": row[0],
                    "document_no": row[1],
                    "vendor_no": row[2],
                    "vendor_name": row[3],
                    "amount": float(row[4]) if row[4] is not None else None,
                    "remaining_amount": float(row[5]) if row[5] is not None else None,
                    "due_date": row[6],
                    "posting_date": row[7],
                    "open": row[8],
                    "currency_code": row[9],
                    "on_hold": row[10],
                }
            return None
        except Exception as e:
            logger.error(f"Fabric SQL exact query failed: {e}")
            return None

    def _query_fuzzy(
        self, conn, entity_code: str, normalized_number: str, col: Dict
    ) -> Optional[Dict]:
        """Query with LIKE for fuzzy match on external_document_no."""
        sql = f"""
            SELECT TOP 1
                {col['external_document_no']},
                {col['document_no']},
                {col['vendor_no']},
                {col['vendor_name']},
                {col['amount']},
                {col['remaining_amount']},
                {col['due_date']},
                {col['posting_date']},
                {col['open']},
                {col['currency_code']},
                {col.get('on_hold', 'on_hold')}
            FROM {self.full_table}
            WHERE {col['entity']} = ?
              AND {col['external_document_no']} LIKE ?
              AND {col['document_type']} = 'Invoice'
            ORDER BY {col['posting_date']} DESC
        """

        try:
            cursor = conn.cursor()
            cursor.execute(sql, (entity_code, f"%{normalized_number}%"))
            row = cursor.fetchone()
            cursor.close()

            if row:
                return {
                    "external_document_no": row[0],
                    "document_no": row[1],
                    "vendor_no": row[2],
                    "vendor_name": row[3],
                    "amount": float(row[4]) if row[4] is not None else None,
                    "remaining_amount": float(row[5]) if row[5] is not None else None,
                    "due_date": row[6],
                    "posting_date": row[7],
                    "open": row[8],
                    "currency_code": row[9],
                    "on_hold": row[10],
                }
            return None
        except Exception as e:
            logger.error(f"Fabric SQL fuzzy query failed: {e}")
            return None

    @staticmethod
    def _normalize_invoice_number(invoice_number: str) -> str:
        """Normalize invoice number for fuzzy matching.

        Strips whitespace, removes common separators (dashes, slashes, spaces).
        """
        normalized = invoice_number.strip()
        # Collapse all whitespace
        normalized = re.sub(r"\s+", "", normalized)
        # Remove dashes and slashes
        normalized = re.sub(r"[-/]", "", normalized)
        return normalized

    def close(self):
        """Close the SQL connection."""
        if self._connection:
            try:
                self._connection.close()
            except Exception:
                pass
            self._connection = None
