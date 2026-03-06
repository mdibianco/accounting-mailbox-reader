"""Microsoft Graph API client for accessing shared mailboxes."""

import json
import logging
import time
from pathlib import Path
from typing import Optional

import msal
import requests

from .config import config

# Retry settings
MAX_RETRIES = 3
RETRY_BACKOFF_BASE = 2  # seconds: 2, 4, 8
RETRYABLE_STATUS_CODES = {429, 500, 502, 503, 504}

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

    def check_health(self, mailbox: str) -> tuple[bool, str]:
        """
        Quick health check: verify token + permissions via a lightweight mail API call.
        Uses _make_request so it benefits from retry logic and token refresh.

        Returns (ok, message).
        """
        result = self._make_request(
            "GET",
            f"/users/{mailbox}/mailFolders/inbox",
            params={"$select": "id"},
        )
        if result is not None:
            return True, "OK"
        return False, "Graph API unreachable - check token, permissions, and network"

    def _make_request(self, method: str, endpoint: str, **kwargs) -> Optional[dict]:
        """Make a request to Graph API with retry on transient errors."""
        self.token = self._get_token()
        if not self.token:
            logger.error("No valid token available")
            return None

        url = f"https://graph.microsoft.com/v1.0{endpoint}"
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json",
        }

        last_error = None
        for attempt in range(MAX_RETRIES):
            try:
                response = requests.request(method, url, headers=headers, timeout=30, **kwargs)

                # Retryable server/throttle errors
                if response.status_code in RETRYABLE_STATUS_CODES:
                    retry_after = int(response.headers.get("Retry-After", RETRY_BACKOFF_BASE ** (attempt + 1)))
                    logger.warning(
                        f"Graph API {response.status_code} (attempt {attempt + 1}/{MAX_RETRIES}), "
                        f"retrying in {retry_after}s..."
                    )
                    time.sleep(retry_after)
                    last_error = f"{response.status_code}: {response.text[:200]}"
                    continue

                # 401/403 = token expired or revoked, clear cache and re-acquire
                if response.status_code in (401, 403) and attempt == 0:
                    logger.warning(
                        f"Token rejected ({response.status_code}), clearing cache and re-authenticating..."
                    )
                    for acct in self.app.get_accounts():
                        self.app.remove_account(acct)
                    self._save_cache()
                    self.token = self._get_token()
                    if self.token:
                        headers["Authorization"] = f"Bearer {self.token}"
                        continue

                response.raise_for_status()

                # Some endpoints return 202/204 with no body
                if response.status_code in (202, 204) or not response.content:
                    return {}
                return response.json()

            except requests.exceptions.Timeout:
                last_error = f"Timeout (attempt {attempt + 1}/{MAX_RETRIES})"
                logger.warning(last_error)
                if attempt < MAX_RETRIES - 1:
                    time.sleep(RETRY_BACKOFF_BASE ** (attempt + 1))
                continue
            except requests.exceptions.ConnectionError as e:
                last_error = f"Connection error (attempt {attempt + 1}/{MAX_RETRIES}): {e}"
                logger.warning(last_error)
                if attempt < MAX_RETRIES - 1:
                    time.sleep(RETRY_BACKOFF_BASE ** (attempt + 1))
                continue
            except requests.exceptions.RequestException as e:
                logger.error(f"Graph API request failed: {e}")
                return None

        logger.error(f"Graph API request failed after {MAX_RETRIES} retries: {last_error}")
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
                "id,webLink,conversationId,from,toRecipients,ccRecipients,subject,bodyPreview,receivedDateTime,"
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

    def forward_message(
        self, mailbox: str, message_id: str, to_addresses: list, comment: str = ""
    ) -> bool:
        """
        Forward a message to one or more recipients.

        Args:
            mailbox: Email address of the mailbox
            message_id: Message ID to forward
            to_addresses: List of email addresses to forward to
            comment: Optional comment to include in the forward

        Returns:
            True if successful, False otherwise
        """
        endpoint = f"/users/{mailbox}/messages/{message_id}/forward"
        payload = {
            "comment": comment,
            "toRecipients": [
                {"emailAddress": {"address": addr}} for addr in to_addresses
            ]
        }
        result = self._make_request("POST", endpoint, json=payload)
        return result is not None

    def create_draft_reply(
        self, mailbox: str, message_id: str, body_html: str
    ) -> Optional[str]:
        """
        Create a draft reply to a message.

        Args:
            mailbox: Email address of the mailbox
            message_id: Message ID to reply to
            body_html: Reply body in HTML format

        Returns:
            Draft message ID if successful, None otherwise
        """
        # First, create the reply draft
        endpoint = f"/users/{mailbox}/messages/{message_id}/createReply"
        draft_result = self._make_request("POST", endpoint, json={})

        if not draft_result or "id" not in draft_result:
            logger.error("Failed to create draft reply")
            return None

        draft_id = draft_result["id"]

        # Now update the draft with the body content
        endpoint = f"/users/{mailbox}/messages/{draft_id}"
        update_result = self._make_request(
            "PATCH",
            endpoint,
            json={
                "body": {
                    "contentType": "HTML",
                    "content": body_html
                }
            }
        )

        if update_result:
            logger.info(f"Draft reply created: {draft_id}")
            return draft_id
        else:
            logger.error(f"Failed to update draft reply body: {draft_id}")
            return None

    def send_mail(
        self, mailbox: str, to_recipients: list, subject: str, body: str, is_html: bool = False
    ) -> bool:
        """
        Send an email message.

        Args:
            mailbox: Email address of the mailbox to send from
            to_recipients: List of recipient email addresses
            subject: Email subject
            body: Email body content
            is_html: Whether the body is HTML or plain text

        Returns:
            True if successful, False otherwise
        """
        endpoint = f"/users/{mailbox}/sendMail"
        payload = {
            "message": {
                "subject": subject,
                "body": {
                    "contentType": "HTML" if is_html else "Text",
                    "content": body
                },
                "toRecipients": [
                    {"emailAddress": {"address": addr}} for addr in to_recipients
                ]
            },
            "saveToSentItems": True
        }
        result = self._make_request("POST", endpoint, json=payload)
        return result is not None

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
        Get ALL messages in Inbox with a specific flag status (with pagination).

        Args:
            mailbox: Email address
            flag_status: "flagged", "complete", or "notFlagged"

        Returns:
            List of all messages or None
        """
        endpoint = f"/users/{mailbox}/mailFolders/inbox/messages"
        all_messages = []
        skip = 0
        page_size = 250

        while True:
            params = {
                "$filter": f"flag/flagStatus eq '{flag_status}'",
                "$select": "id,subject,receivedDateTime,flag,categories",
                "$top": page_size,
                "$skip": skip,
            }
            result = self._make_request("GET", endpoint, params=params)
            if result and "value" in result:
                messages = result["value"]
                if not messages:
                    break
                all_messages.extend(messages)
                skip += page_size
            else:
                break

        return all_messages if all_messages else None

    def get_folder_id(
        self, mailbox: str, folder_path: str
    ) -> Optional[str]:
        """
        Find a mail folder by path (e.g., 'Reminders' or 'Parent/Child').
        Read-only — does not create missing folders.

        Returns folder ID or None if not found.
        """
        parts = [p.strip() for p in folder_path.split("/") if p.strip()]
        parent_id = None

        for part in parts:
            if parent_id:
                endpoint = f"/users/{mailbox}/mailFolders/{parent_id}/childFolders"
            else:
                endpoint = f"/users/{mailbox}/mailFolders"

            result = self._make_request("GET", endpoint, params={
                "$filter": f"displayName eq '{part}'",
                "$top": 1,
            })

            if result and result.get("value"):
                parent_id = result["value"][0]["id"]
            else:
                logger.error(f"Folder not found: {part}")
                return None

        return parent_id

    def get_folder_messages(
        self, mailbox: str, folder_id: str,
        max_results: int = 1000, days_back: int = 60,
    ) -> Optional[list]:
        """
        Get messages from a specific folder with pagination support.

        Returns list of messages (up to max_results) or None.
        """
        from datetime import datetime, timedelta

        days_ago = datetime.utcnow() - timedelta(days=days_back)
        date_filter = f"receivedDateTime ge {days_ago.isoformat()}Z"

        endpoint = f"/users/{mailbox}/mailFolders/{folder_id}/messages"
        params = {
            "$filter": date_filter,
            "$top": min(max_results, 250),
            "$orderby": "receivedDateTime asc",
            "$select": (
                "id,webLink,conversationId,from,toRecipients,ccRecipients,subject,bodyPreview,receivedDateTime,"
                "hasAttachments,importance,isRead,categories"
            ),
        }

        all_messages = []
        result = self._make_request("GET", endpoint, params=params)

        while result:
            messages = result.get("value", [])
            all_messages.extend(messages)

            if len(all_messages) >= max_results:
                break

            next_link = result.get("@odata.nextLink")
            if not next_link:
                break

            # Follow pagination
            self.token = self._get_token()
            if not self.token:
                break
            try:
                response = requests.get(next_link, headers={
                    "Authorization": f"Bearer {self.token}",
                    "Content-Type": "application/json",
                })
                response.raise_for_status()
                result = response.json()
            except requests.exceptions.RequestException as e:
                logger.error(f"Pagination failed: {e}")
                break

        return all_messages[:max_results] if all_messages else None

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
