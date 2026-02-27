"""Microsoft Graph API client for accessing shared mailboxes."""

import json
import logging
from pathlib import Path
from typing import Optional

import msal
import requests

from .config import config

logger = logging.getLogger(__name__)


class GraphAPIClient:
    """Client for Microsoft Graph API."""

    def __init__(self):
        """Initialize Graph API client."""
        self.client_id = config.azure_client_id
        self.tenant_id = config.azure_tenant_id
        self.client_secret = config.azure_client_secret
        self.scopes = ["https://graph.microsoft.com/.default"]

        # Set up token cache
        self.cache_file = Path.home() / ".accounting_mailbox_reader" / "token_cache.bin"
        self.cache_file.parent.mkdir(parents=True, exist_ok=True)
        self.token_cache = msal.SerializableTokenCache()

        # Load existing cache if available
        if self.cache_file.exists():
            with open(self.cache_file, "r") as f:
                self.token_cache.deserialize(f.read())

        self.app = None
        self.token = None
        self._initialize_auth()

    def _initialize_auth(self):
        """Initialize MSAL authentication."""
        # Try Client Secret first (non-interactive, best for automation)
        if self.client_secret and self.client_secret.strip():
            logger.info("Using Client Secret authentication (non-interactive)...")
            self.app = msal.ConfidentialClientApplication(
                client_id=self.client_id,
                client_credential=self.client_secret,
                authority=f"https://login.microsoftonline.com/{self.tenant_id}",
                token_cache=self.token_cache
            )
            logger.info("Client Secret authentication initialized")
        else:
            logger.warning(
                "Client Secret not found in environment. "
                "Using Device Code Flow (interactive login with persistent cache)."
            )
            self.app = msal.PublicClientApplication(
                client_id=self.client_id,
                authority=f"https://login.microsoftonline.com/{self.tenant_id}",
                token_cache=self.token_cache
            )
            logger.info("Device Code Flow authentication initialized with persistent cache")

    def _save_cache(self):
        """Save token cache to disk."""
        if self.token_cache.has_state_changed:
            with open(self.cache_file, "w") as f:
                f.write(self.token_cache.serialize())
            logger.debug(f"Token cache saved to {self.cache_file}")

    def _get_token(self) -> Optional[str]:
        """Get access token for Graph API with caching."""
        # First, try to get token from cache
        accounts = self.app.get_accounts()
        if accounts:
            logger.debug(f"Found {len(accounts)} cached account(s)")
            result = self.app.acquire_token_silent(self.scopes, account=accounts[0])
            if result and "access_token" in result:
                logger.debug("Using cached token (no login required)")
                self._save_cache()
                return result["access_token"]

        # No cached token, need to acquire new one
        if self.client_secret and self.client_secret.strip():
            # Client secret flow (non-interactive)
            logger.info("Acquiring new token with client secret...")
            result = self.app.acquire_token_for_client(scopes=self.scopes)
        else:
            # Device code flow (interactive)
            logger.info("No cached token found. Starting device code flow...")
            flow = self.app.initiate_device_flow(scopes=self.scopes)

            if "user_code" not in flow:
                logger.error("Failed to create device flow")
                return None

            # Display device code to user
            print(f"\n{'='*70}")
            print("To sign in, use a web browser to open the page:")
            print(f"  {flow['verification_uri']}")
            print(f"\nEnter this code:")
            print(f"  {flow['user_code']}")
            print(f"{'='*70}\n")

            result = self.app.acquire_token_by_device_flow(flow)

        if result and "access_token" in result:
            logger.info("✓ Successfully acquired access token")
            self._save_cache()
            return result["access_token"]
        else:
            error = result.get("error_description", result.get("error", "Unknown error"))
            logger.error(f"Failed to acquire token: {error}")
            return None

    def _make_request(self, method: str, endpoint: str, **kwargs) -> Optional[dict]:
        """Make a request to Graph API."""
        self.token = self._get_token()
        if not self.token:
            logger.error("No valid token available")
            return None

        url = f"https://graph.microsoft.com/v1.0{endpoint}"
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json",
        }

        try:
            response = requests.request(method, url, headers=headers, **kwargs)
            response.raise_for_status()
            return response.json()
        except requests.exceptions.RequestException as e:
            logger.error(f"Graph API request failed: {e}")
            return None

    def get_mailbox_messages(
        self,
        mailbox: str,
        max_results: int = 50,
        days_back: int = 7,
        search_query: Optional[str] = None,
    ) -> Optional[list]:
        """
        Get messages from a shared mailbox.

        Args:
            mailbox: Email address of the mailbox
            max_results: Maximum number of messages to return
            days_back: How many days back to search
            search_query: Optional search query

        Returns:
            List of messages or None if error
        """
        # Build filter for date range
        from datetime import datetime, timedelta

        days_ago = datetime.utcnow() - timedelta(days=days_back)
        date_filter = f"receivedDateTime ge {days_ago.isoformat()}Z"

        # Build OData query
        filter_query = date_filter
        if search_query:
            filter_query += f" and ({search_query})"

        endpoint = f"/users/{mailbox}/mailFolders/inbox/messages"
        params = {
            "$filter": filter_query,
            "$top": min(max_results, 250),  # Graph API limit
            "$orderby": "receivedDateTime desc",
            "$select": (
                "id,from,subject,bodyPreview,receivedDateTime,"
                "hasAttachments,importance,isRead"
            ),
        }

        result = self._make_request("GET", endpoint, params=params)
        if result and "value" in result:
            return result["value"]
        return None

    def get_message_body(self, mailbox: str, message_id: str) -> Optional[str]:
        """Get the full body of a message."""
        endpoint = f"/users/{mailbox}/messages/{message_id}"
        params = {"$select": "body"}

        result = self._make_request("GET", endpoint, params=params)
        if result and "body" in result:
            return result["body"].get("content", "")
        return None

    def get_message_attachments(
        self, mailbox: str, message_id: str
    ) -> Optional[list]:
        """Get attachments metadata for a message."""
        endpoint = f"/users/{mailbox}/messages/{message_id}/attachments"
        params = {
            "$select": "id,name,contentType,size",
        }

        result = self._make_request("GET", endpoint, params=params)
        if result and "value" in result:
            return result["value"]
        return None

    def get_attachment_content(
        self, mailbox: str, message_id: str, attachment_id: str
    ) -> Optional[bytes]:
        """Get the binary content of an attachment."""
        self.token = self._get_token()
        if not self.token:
            return None

        url = (
            f"https://graph.microsoft.com/v1.0/users/{mailbox}/"
            f"messages/{message_id}/attachments/{attachment_id}/$value"
        )
        headers = {"Authorization": f"Bearer {self.token}"}

        try:
            response = requests.get(url, headers=headers)
            response.raise_for_status()
            return response.content
        except requests.exceptions.RequestException as e:
            logger.error(f"Failed to download attachment: {e}")
            return None

    def update_message(
        self, mailbox: str, message_id: str, updates: dict
    ) -> Optional[dict]:
        """
        Update properties on a mailbox message (e.g., categories).

        Args:
            mailbox: Email address of the mailbox
            message_id: Message ID
            updates: Dict of properties to update (e.g., {"categories": ["VEN-REM"]})

        Returns:
            Updated message dict or None on error
        """
        endpoint = f"/users/{mailbox}/messages/{message_id}"
        return self._make_request("PATCH", endpoint, json=updates)

    def flag_message(
        self, mailbox: str, message_id: str, flag_status: str = "flagged"
    ) -> Optional[dict]:
        """
        Set the flag status on a message.

        Args:
            mailbox: Email address of the mailbox
            message_id: Message ID
            flag_status: "flagged", "complete", or "notFlagged"
        """
        endpoint = f"/users/{mailbox}/messages/{message_id}"
        return self._make_request("PATCH", endpoint, json={
            "flag": {"flagStatus": flag_status}
        })

    def move_message(
        self, mailbox: str, message_id: str, destination_folder_id: str
    ) -> Optional[dict]:
        """
        Move a message to a different folder.

        Returns the moved message (with new ID) or None on error.
        """
        endpoint = f"/users/{mailbox}/messages/{message_id}/move"
        return self._make_request("POST", endpoint, json={
            "destinationId": destination_folder_id
        })

    def get_or_create_folder(
        self, mailbox: str, folder_path: str
    ) -> Optional[str]:
        """
        Get or create a nested mail folder path (e.g., "ARCHIVE/PROCESSED BY AGENT").

        Returns the folder ID of the deepest folder, or None on error.
        """
        parts = [p.strip() for p in folder_path.split("/") if p.strip()]
        parent_id = None  # None = top-level (under mailbox root)

        for part in parts:
            # List child folders of current parent
            if parent_id:
                endpoint = f"/users/{mailbox}/mailFolders/{parent_id}/childFolders"
            else:
                endpoint = f"/users/{mailbox}/mailFolders"

            result = self._make_request("GET", endpoint, params={
                "$filter": f"displayName eq '{part}'",
                "$top": 1,
            })

            if result and result.get("value"):
                # Folder exists
                parent_id = result["value"][0]["id"]
            else:
                # Create the folder
                if parent_id:
                    create_endpoint = f"/users/{mailbox}/mailFolders/{parent_id}/childFolders"
                else:
                    create_endpoint = f"/users/{mailbox}/mailFolders"

                created = self._make_request("POST", create_endpoint, json={
                    "displayName": part
                })
                if not created:
                    logger.error(f"Failed to create folder: {part}")
                    return None
                parent_id = created["id"]
                logger.info(f"Created mail folder: {part} (id: {parent_id})")

        return parent_id

    def get_inbox_messages_by_flag(
        self, mailbox: str, flag_status: str = "complete"
    ) -> Optional[list]:
        """
        Get messages in Inbox with a specific flag status.

        Args:
            mailbox: Email address
            flag_status: "flagged", "complete", or "notFlagged"

        Returns:
            List of messages or None
        """
        endpoint = f"/users/{mailbox}/mailFolders/inbox/messages"
        params = {
            "$filter": f"flag/flagStatus eq '{flag_status}'",
            "$select": "id,subject,receivedDateTime,flag,categories",
            "$top": 250,
        }
        result = self._make_request("GET", endpoint, params=params)
        if result and "value" in result:
            return result["value"]
        return None

    def get_sharepoint_site(self, site_hostname: str, site_path: str) -> Optional[dict]:
        """
        Get SharePoint site information.

        Args:
            site_hostname: SharePoint hostname (e.g., 'planted.sharepoint.com')
            site_path: Site path (e.g., '/sites/finance')

        Returns:
            Site information dict or None
        """
        endpoint = f"/sites/{site_hostname}:/{site_path.lstrip('/')}"
        result = self._make_request("GET", endpoint)
        if result and "id" in result:
            return result
        return None

    def get_sharepoint_drive(
        self, site_hostname: str, site_path: str, drive_name: str = "Documents"
    ) -> Optional[dict]:
        """
        Get SharePoint drive information.

        Args:
            site_hostname: SharePoint hostname
            site_path: Site path
            drive_name: Drive name (default: Documents)

        Returns:
            Drive information dict or None
        """
        site = self.get_sharepoint_site(site_hostname, site_path)
        if not site:
            logger.error(f"Could not get SharePoint site: {site_hostname}{site_path}")
            return None

        logger.debug(f"Found site: {site.get('displayName', 'Unknown')} (id: {site.get('id')})")

        # Try to get the default document library first
        endpoint = f"/sites/{site['id']}/drive"
        result = self._make_request("GET", endpoint)
        if result and result.get("name"):
            logger.info(f"Found default drive: {result.get('name')}")
            # Check if this is the drive we're looking for
            if result.get("name", "").lower() == drive_name.lower():
                return result

            # Save default drive for fallback
            default_drive = result

        # If default drive doesn't match, get all drives
        endpoint = f"/sites/{site['id']}/drives"
        result = self._make_request("GET", endpoint)

        if not result or "value" not in result:
            logger.error(f"Failed to get drives list. Response: {result}")
            # Try to use default drive if we found one
            if 'default_drive' in locals():
                logger.warning(f"Using default drive: {default_drive.get('name')}")
                return default_drive
            return None

        drives = result.get("value", [])
        logger.info(f"Found {len(drives)} drive(s) on site")

        # Log all available drives for debugging
        for drive in drives:
            logger.info(f"  - Drive: {drive.get('name')} (type: {drive.get('driveType')})")

        # Find drive by name (case-insensitive)
        for drive in drives:
            if drive.get("name", "").lower() == drive_name.lower():
                logger.info(f"Matched drive: {drive.get('name')}")
                return drive

        # If not found, log available drives
        available_drives = [d.get("name") for d in drives]
        logger.error(
            f"Could not find drive: {drive_name}. "
            f"Available drives: {', '.join(available_drives) if available_drives else 'NONE'}"
        )

        # Try to use default drive as fallback
        if 'default_drive' in locals():
            logger.warning(f"Falling back to default drive: {default_drive.get('name')}")
            return default_drive

        return None

    def upload_to_sharepoint(
        self,
        site_hostname: str,
        site_path: str,
        drive_name: str,
        folder_path: str,
        file_name: str,
        file_content: bytes,
    ) -> Optional[dict]:
        """
        Upload a file to SharePoint.

        Args:
            site_hostname: SharePoint hostname
            site_path: Site path
            drive_name: Drive name
            folder_path: Path within drive (e.g., 'Accounting/Mailbox/emails')
            file_name: Name of file to create
            file_content: Binary content of file

        Returns:
            Upload result dict or None
        """
        site = self.get_sharepoint_site(site_hostname, site_path)
        if not site:
            logger.error(f"Could not find SharePoint site: {site_path}")
            return None

        drive = self.get_sharepoint_drive(site_hostname, site_path, drive_name)
        if not drive:
            logger.error(f"Could not find drive: {drive_name}")
            return None

        self.token = self._get_token()
        if not self.token:
            return None

        # Construct the upload path
        # /sites/{site-id}/drive/root:/{item-path}:/content
        upload_path = f"/sites/{site['id']}/drive/root:/{folder_path}/{file_name}:/content"

        url = f"https://graph.microsoft.com/v1.0{upload_path}"
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/octet-stream",
        }

        try:
            response = requests.put(url, headers=headers, data=file_content)
            response.raise_for_status()
            logger.info(f"Successfully uploaded {file_name} to SharePoint")
            return response.json()
        except requests.exceptions.RequestException as e:
            logger.error(f"Failed to upload to SharePoint: {e}")
            return None
