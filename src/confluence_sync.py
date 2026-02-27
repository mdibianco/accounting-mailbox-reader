"""Sync categories from Confluence."""

import json
import logging
from pathlib import Path
from datetime import datetime
from typing import Optional, List, Dict
import requests
from bs4 import BeautifulSoup

from .config import config

logger = logging.getLogger(__name__)


class ConfluenceSyncer:
    """Syncs category table from Confluence."""

    def __init__(self):
        """Initialize Confluence client."""
        self.confluence_url = config.confluence_page_url
        self.email = config.confluence_email
        self.api_token = config.confluence_api_token
        self.cache_file = Path(__file__).parent.parent / "data" / "categories_cache.json"

        if not self.email or not self.api_token:
            logger.warning(
                "Confluence credentials not configured. "
                "Set CONFLUENCE_EMAIL and CONFLUENCE_API_TOKEN in .env"
            )

    def sync_categories(self) -> bool:
        """
        Fetch categories from Confluence and update cache.

        Returns:
            True if sync successful, False otherwise
        """
        if not self.email or not self.api_token:
            logger.error("Cannot sync: Confluence credentials not configured")
            return False

        try:
            logger.info(f"Syncing categories from Confluence: {self.confluence_url}")

            # Fetch Confluence page content
            content = self._fetch_page_content()
            if not content:
                logger.error("Failed to fetch Confluence page content")
                return False

            # Parse categories from HTML
            categories = self._parse_categories_from_html(content)
            if not categories:
                logger.error("No categories found in Confluence page")
                return False

            # Save to cache
            self._save_cache(categories)

            logger.info(f"✓ Successfully synced {len(categories)} categories from Confluence")
            return True

        except Exception as e:
            logger.error(f"Failed to sync categories: {e}", exc_info=True)
            return False

    def _fetch_page_content(self) -> Optional[str]:
        """Fetch Confluence page HTML content."""
        try:
            # Extract page ID from URL
            # URL format: https://eatplanted.atlassian.net/wiki/spaces/FC/pages/2361458692/...
            page_id = self.confluence_url.split("/pages/")[1].split("/")[0]

            # Construct API URL
            api_url = f"https://eatplanted.atlassian.net/wiki/rest/api/content/{page_id}?expand=body.storage"

            # Make request with basic auth
            response = requests.get(
                api_url,
                auth=(self.email, self.api_token),
                headers={"Accept": "application/json"}
            )
            response.raise_for_status()

            data = response.json()
            return data["body"]["storage"]["value"]

        except Exception as e:
            logger.error(f"Failed to fetch Confluence page: {e}")
            return None

    def _parse_categories_from_html(self, html_content: str) -> List[Dict]:
        """
        Parse categories from Confluence HTML table.

        Expected table structure:
        | Category ID | Category Name | Description | Expected Action | Key Indicators | Multi-Language Keywords | Exclusions |
        """
        try:
            soup = BeautifulSoup(html_content, 'html.parser')

            # Find the first table in the page
            table = soup.find('table')
            if not table:
                logger.error("No table found in Confluence page")
                return []

            # Extract rows (skip header row)
            rows = table.find_all('tr')[1:]  # Skip header

            categories = []
            for row in rows:
                cells = row.find_all(['td', 'th'])
                if len(cells) < 7:
                    logger.warning(f"Skipping row with insufficient columns: {len(cells)}")
                    continue

                # Extract cell text
                category_id = cells[0].get_text(strip=True)
                category_name = cells[1].get_text(strip=True)
                description = cells[2].get_text(strip=True)
                expected_action = cells[3].get_text(strip=True)
                key_indicators = cells[4].get_text(strip=True)
                keywords_raw = cells[5].get_text(strip=True)
                exclusions = cells[6].get_text(strip=True)

                # Parse multi-language keywords
                keywords = self._parse_keywords(keywords_raw)

                category = {
                    "id": category_id,
                    "name": category_name,
                    "description": description,
                    "expected_action": expected_action,
                    "key_indicators": key_indicators,
                    "keywords": keywords,
                    "exclusions": exclusions
                }

                categories.append(category)
                logger.debug(f"Parsed category: {category_id} - {category_name}")

            return categories

        except Exception as e:
            logger.error(f"Failed to parse categories from HTML: {e}", exc_info=True)
            return []

    def _parse_keywords(self, keywords_raw: str) -> Dict[str, List[str]]:
        """
        Parse multi-language keywords from text.

        Expected format:
        EN: keyword1, keyword2
        DE: keyword1, keyword2
        FR: keyword1, keyword2
        IT: keyword1, keyword2
        """
        keywords = {"en": [], "de": [], "fr": [], "it": []}

        try:
            # Split by language markers
            for line in keywords_raw.split('\n'):
                line = line.strip()
                if not line:
                    continue

                # Check for language markers
                if line.startswith('EN:') or line.startswith('**EN:**'):
                    keywords['en'] = [k.strip() for k in line.split(':', 1)[1].split(',') if k.strip()]
                elif line.startswith('DE:') or line.startswith('**DE:**'):
                    keywords['de'] = [k.strip() for k in line.split(':', 1)[1].split(',') if k.strip()]
                elif line.startswith('FR:') or line.startswith('**FR:**'):
                    keywords['fr'] = [k.strip() for k in line.split(':', 1)[1].split(',') if k.strip()]
                elif line.startswith('IT:') or line.startswith('**IT:**'):
                    keywords['it'] = [k.strip() for k in line.split(':', 1)[1].split(',') if k.strip()]

        except Exception as e:
            logger.warning(f"Failed to parse keywords: {e}")

        return keywords

    def _save_cache(self, categories: List[Dict]):
        """Save categories to local cache file."""
        try:
            self.cache_file.parent.mkdir(parents=True, exist_ok=True)

            data = {
                "last_synced": datetime.utcnow().isoformat() + "Z",
                "confluence_page_url": self.confluence_url,
                "categories": categories
            }

            with open(self.cache_file, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)

            logger.info(f"Categories cache saved to: {self.cache_file}")

        except Exception as e:
            logger.error(f"Failed to save categories cache: {e}")
            raise
