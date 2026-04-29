"""Dataverse Web API client for managing Knowledge Articles.

Creates knowledge articles and transitions them to Published state.
"""

import logging
import time
import uuid
from datetime import datetime, timezone
import requests
from kb_loader.auth import AuthClient
from kb_loader.config import Config

logger = logging.getLogger(__name__)

MAX_RETRIES = 3
RETRY_BASE_DELAY = 2

# Knowledge Article state/status codes (standard D365 Customer Service)
# These are the OOB values; override via config if your org differs.
STATE_DRAFT = 0
STATUS_DRAFT = 2
STATE_APPROVED = 1
STATUS_APPROVED = 5
STATE_PUBLISHED = 3
STATUS_PUBLISHED = 7

# English language locale ID
ENGLISH_LOCALE_ID = 1033


class DataverseClient:
    """Client for creating and publishing Knowledge Articles in Dataverse."""

    def __init__(self, auth: AuthClient, config: Config):
        self.auth = auth
        self.config = config
        self.session = requests.Session()
        self._language_id: str | None = None

    def _headers(self) -> dict:
        return {
            "Authorization": f"Bearer {self.auth.get_dataverse_token(self.config.dataverse_url)}",
            "Accept": "application/json",
            "Content-Type": "application/json",
            "OData-MaxVersion": "4.0",
            "OData-Version": "4.0",
        }

    def _api(self, path: str) -> str:
        return f"{self.config.dataverse_api_url}/{path.lstrip('/')}"

    def _request(self, method: str, url: str, **kwargs) -> requests.Response:
        """Make an HTTP request with retry logic."""
        for attempt in range(MAX_RETRIES):
            resp = self.session.request(method, url, headers=self._headers(), **kwargs)
            if resp.status_code == 429:
                retry_after = int(resp.headers.get("Retry-After", RETRY_BASE_DELAY * (attempt + 1)))
                logger.warning(f"Throttled by Dataverse. Waiting {retry_after}s...")
                time.sleep(retry_after)
                continue
            if resp.status_code >= 500:
                delay = RETRY_BASE_DELAY * (attempt + 1)
                logger.warning(f"Server error {resp.status_code}. Retrying in {delay}s...")
                time.sleep(delay)
                continue
            return resp
        raise RuntimeError(f"Dataverse request failed after {MAX_RETRIES} retries: {url}")

    def _get_language_id(self) -> str:
        """Look up the knowledgearticle language record for English (1033)."""
        if self._language_id:
            return self._language_id

        url = self._api(
            f"languagelocales?$filter=localeid eq {ENGLISH_LOCALE_ID}&$select=languagelocaleid"
        )
        resp = self._request("GET", url)
        resp.raise_for_status()
        data = resp.json()
        records = data.get("value", [])

        if not records:
            raise RuntimeError(
                f"Language locale {ENGLISH_LOCALE_ID} not found in Dataverse. "
                "Ensure the language pack is installed."
            )

        self._language_id = records[0]["languagelocaleid"]
        logger.info(f"Resolved English language locale ID: {self._language_id}")
        return self._language_id

    def find_existing_article(self, title: str) -> dict | None:
        """Check if a knowledge article with the given title already exists."""
        # Escape single quotes in title for OData filter
        safe_title = title.replace("'", "''")
        url = self._api(
            f"knowledgearticles?$filter=title eq '{safe_title}'"
            f"&$select=knowledgearticleid,title,statecode,statuscode"
            f"&$top=1"
        )
        resp = self._request("GET", url)
        resp.raise_for_status()
        records = resp.json().get("value", [])
        return records[0] if records else None

    def create_article(
        self,
        title: str,
        html_content: str,
        source_path: str = "",
    ) -> str:
        """Create a new Knowledge Article in Draft state.

        Args:
            title: Article title (from filename).
            html_content: HTML body content.
            source_path: Original SharePoint path for traceability.

        Returns:
            The knowledgearticleid of the created record.
        """
        language_id = self._get_language_id()

        article_data = {
            "title": title,
            "content": html_content,
            "keywords": source_path,
            "description": f"Auto-imported from SharePoint: {source_path}",
            "languagelocaleid@odata.bind": f"/languagelocales({language_id})",
            "isrootarticle": False,
            "createdon": datetime.now(timezone.utc).isoformat(),
            # Manual creation mode
            "msdyn_creationmode": 0,  # 0 = Manual
        }

        url = self._api("knowledgearticles")
        resp = self._request("POST", url, json=article_data)

        if resp.status_code not in (200, 201, 204):
            error_body = resp.text
            raise RuntimeError(
                f"Failed to create article '{title}': {resp.status_code} - {error_body}"
            )

        # Extract the article ID from the response
        if resp.status_code == 204:
            # ID is in the OData-EntityId header
            entity_id = resp.headers.get("OData-EntityId", "")
            article_id = entity_id.split("(")[-1].rstrip(")")
        else:
            article_id = resp.json().get("knowledgearticleid", "")

        logger.info(f"Created article '{title}' with ID: {article_id}")
        return article_id

    def update_article_content(self, article_id: str, title: str, html_content: str, source_path: str = ""):
        """Update an existing Knowledge Article's content."""
        language_id = self._get_language_id()

        article_data = {
            "title": title,
            "content": html_content,
            "keywords": source_path,
            "description": f"Auto-imported from SharePoint: {source_path}",
        }

        url = self._api(f"knowledgearticles({article_id})")
        resp = self._request("PATCH", url, json=article_data)

        if resp.status_code not in (200, 204):
            raise RuntimeError(
                f"Failed to update article '{title}': {resp.status_code} - {resp.text}"
            )

        logger.info(f"Updated article '{title}' ({article_id})")

    def publish_article(self, article_id: str):
        """Transition a Knowledge Article from Draft → Published.

        Uses the SetState approach via PATCH. If the org requires intermediate
        transitions (Draft → Approved → Published), this handles both steps.
        """
        # Step 1: Try direct transition to Published
        try:
            self._set_state(article_id, STATE_PUBLISHED, STATUS_PUBLISHED)
            logger.info(f"Article {article_id} published directly.")
            return
        except RuntimeError as e:
            logger.info(f"Direct publish failed, trying via Approved state: {e}")

        # Step 2: Transition to Approved first, then Published
        self._set_state(article_id, STATE_APPROVED, STATUS_APPROVED)
        logger.info(f"Article {article_id} approved.")
        self._set_state(article_id, STATE_PUBLISHED, STATUS_PUBLISHED)
        logger.info(f"Article {article_id} published.")

    def _set_state(self, article_id: str, statecode: int, statuscode: int):
        """Update the state and status of a knowledge article."""
        url = self._api(f"knowledgearticles({article_id})")
        payload = {
            "statecode": statecode,
            "statuscode": statuscode,
        }
        resp = self._request("PATCH", url, json=payload)

        if resp.status_code not in (200, 204):
            raise RuntimeError(
                f"Failed to set state ({statecode}/{statuscode}) for article {article_id}: "
                f"{resp.status_code} - {resp.text}"
            )
