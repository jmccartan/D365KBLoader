# =====================================================================
# D365 Knowledge Base Loader
# Copyright (c) 2026. All rights reserved.
# Licensed under the MIT License. See the LICENSE file in the project
# root for the full text.
# =====================================================================

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

# D365 knowledgearticle entity field length limits.
# These are the standard out-of-box maximums; if a custom org has raised
# them the truncation here is still safe (we just send a shorter value).
TITLE_MAX_LENGTH = 1000
DESCRIPTION_MAX_LENGTH = 155
KEYWORDS_MAX_LENGTH = 100


def _truncate(text, max_length: int) -> str:
    """Truncate `text` to `max_length` characters, appending an ellipsis if shortened.

    Defensively coerces None / non-string values to an empty string so the
    payload never contains JSON null for a string field (Dataverse rejects null
    on most string attributes).
    """
    if text is None:
        return ""
    if not isinstance(text, str):
        text = str(text)
    if len(text) <= max_length:
        return text
    if max_length <= 3:
        return text[:max_length]
    return text[: max_length - 3] + "..."


def validate_article_payload(
    title: str,
    source_path: str,
    has_content: bool,
) -> list[tuple[str, str]]:
    """Pre-flight checks on values that will be sent to D365.

    Returns a list of (severity, message) tuples. Severity is one of:
      - "error":   the article cannot be created (D365 will reject it)
      - "warning": the article will be created but a value will be silently
                   truncated to fit a D365 field-length limit
      - "info":    minor notice (e.g. empty content body)

    Empty list means the payload is clean.
    """
    issues: list[tuple[str, str]] = []

    # Title
    if not title or not title.strip():
        issues.append((
            "error",
            "Title is empty (the file name has no usable stem). "
            "D365 will reject this article — rename the source file.",
        ))
    elif len(title) > TITLE_MAX_LENGTH:
        truncated = _truncate(title, TITLE_MAX_LENGTH)
        issues.append((
            "warning",
            f"Title is {len(title)} chars (D365 max is {TITLE_MAX_LENGTH}). "
            f"Will be saved as: {truncated!r}",
        ))

    # Description (auto-generated from source_path)
    desc = f"Auto-imported from SharePoint: {source_path}"
    if len(desc) > DESCRIPTION_MAX_LENGTH:
        issues.append((
            "warning",
            f"Description ({len(desc)} chars) exceeds D365 max of "
            f"{DESCRIPTION_MAX_LENGTH}; will be truncated with an ellipsis.",
        ))

    # Keywords (= source path)
    if source_path and len(source_path) > KEYWORDS_MAX_LENGTH:
        issues.append((
            "warning",
            f"Keywords/source-path ({len(source_path)} chars) exceeds D365 max "
            f"of {KEYWORDS_MAX_LENGTH}; will be truncated with an ellipsis.",
        ))

    # Body content
    if not has_content:
        issues.append((
            "info",
            "Document converted to empty HTML — article would be created "
            "with an empty body. Consider checking the source document.",
        ))

    return issues


class DataverseClient:
    """Client for creating and publishing Knowledge Articles in Dataverse."""

    def __init__(self, auth: AuthClient, config: Config):
        self.auth = auth
        self.config = config
        self.session = requests.Session()
        self._language_id: str | None = None
        # Cached Business Process Flow metadata: (workflowid, instance_entity_set, {stage_name: stage_id})
        # `False` = lookup attempted but no suitable BPF was found, so don't keep trying.
        self._kb_bpf: tuple[str, str, dict[str, str]] | None | bool = None

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
            try:
                resp = self.session.request(method, url, headers=self._headers(), **kwargs)
            except requests.exceptions.ConnectionError as e:
                # DNS failures, refused connections, no-route errors all surface as
                # ConnectionError. Translate to a friendly message so the user
                # knows it's a URL/network issue, not a bug.
                msg = str(e).lower()
                if "nameresolutionerror" in msg or "getaddrinfo failed" in msg or "name or service not known" in msg:
                    host = self.config.dataverse_url.replace("https://", "").rstrip("/")
                    raise RuntimeError(
                        f"Couldn't reach the Dataverse server: {host}\n\n"
                        "The hostname doesn't exist in DNS. Possible causes:\n"
                        "  - Typo in the Dataverse URL (check Settings)\n"
                        "  - The D365 environment has been deleted or renamed\n"
                        "  - You're offline or behind a VPN that blocks DNS\n\n"
                        "Verify the URL in the Power Platform admin center "
                        "(https://admin.powerplatform.microsoft.com)."
                    )
                # Other connection errors — retry on the first attempt, then surface
                if attempt < MAX_RETRIES - 1:
                    delay = RETRY_BASE_DELAY * (attempt + 1)
                    logger.warning(f"Connection error: {e}. Retrying in {delay}s...")
                    time.sleep(delay)
                    continue
                raise RuntimeError(f"Could not connect to Dataverse: {e}")

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
            f"languagelocale?$filter=localeid eq {ENGLISH_LOCALE_ID}&$select=languagelocaleid"
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
        # Escape single quotes in title for OData filter.
        # Other special characters (& # + etc.) are handled by passing the
        # filter via the `params` arg, which URL-encodes the whole value.
        safe_title = title.replace("'", "''")
        url = self._api("knowledgearticles")
        resp = self._request(
            "GET",
            url,
            params={
                "$filter": f"title eq '{safe_title}'",
                "$select": "knowledgearticleid,title,statecode,statuscode",
                "$top": "1",
            },
        )
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
            "title": _truncate(title, TITLE_MAX_LENGTH),
            "content": html_content,
            "keywords": _truncate(source_path, KEYWORDS_MAX_LENGTH),
            "description": _truncate(
                f"Auto-imported from SharePoint: {source_path}",
                DESCRIPTION_MAX_LENGTH,
            ),
            "languagelocaleid@odata.bind": f"/languagelocale({language_id})",
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
            "title": _truncate(title, TITLE_MAX_LENGTH),
            "content": html_content,
            "keywords": _truncate(source_path, KEYWORDS_MAX_LENGTH),
            "description": _truncate(
                f"Auto-imported from SharePoint: {source_path}",
                DESCRIPTION_MAX_LENGTH,
            ),
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
        After the state transition, also advances the Business Process Flow
        widget on the form (Author → Review → Publish) on a best-effort basis
        — failures there do not fail the publish.
        """
        # Step 1: Try direct transition to Published
        try:
            self._set_state(article_id, STATE_PUBLISHED, STATUS_PUBLISHED)
            logger.info(f"Article {article_id} published directly.")
        except RuntimeError as e:
            logger.info(f"Direct publish failed, trying via Approved state: {e}")
            # Step 2: Transition to Approved first, then Published
            self._set_state(article_id, STATE_APPROVED, STATUS_APPROVED)
            logger.info(f"Article {article_id} approved.")
            self._set_state(article_id, STATE_PUBLISHED, STATUS_PUBLISHED)
            logger.info(f"Article {article_id} published.")

        # Step 3: Advance the BPF widget on the form (cosmetic, best-effort)
        self.advance_bpf_to_publish_stage(article_id)

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

    # ── Business Process Flow advancement ──────────────────────────────
    # The Knowledge Article form in D365 Customer Service shows a BPF widget
    # at the top with stages like Author → Review → Publish. The BPF state is
    # tracked on a separate auto-generated entity (e.g. `newprocess`), not on
    # the article itself. After we set statuscode=Published, the BPF widget
    # would still show "Author" until we also advance the BPF instance.
    #
    # The methods below resolve the BPF metadata once (cached), then either
    # update an existing BPF instance row for the article or create a new one,
    # setting `activestageid` to the final "Publish" stage and `traversedpath`
    # to the natural Author→Review→Publish breadcrumb when those stage names
    # exist. Any failure is logged and swallowed so the load does not fail.

    def _get_kb_bpf_metadata(self) -> tuple[str, str, dict[str, str]] | None:
        """Resolve & cache the Knowledge Article BPF.

        Returns (workflowid, instance_entity_set, stage_name_to_id). Returns
        None if no suitable active BPF exists (BPF advancement is skipped).
        """
        if self._kb_bpf is False:
            return None
        if isinstance(self._kb_bpf, tuple):
            return self._kb_bpf

        try:
            r = self._request(
                "GET",
                self._api("workflows"),
                params={
                    # category 4 = Business Process Flow; statecode 1 = Activated
                    "$filter": "primaryentity eq 'knowledgearticle' and category eq 4 and statecode eq 1",
                    "$select": "workflowid,name,uniquename",
                },
            )
            r.raise_for_status()
            bpfs = r.json().get("value", [])
            # Prefer the OOB "New Process" BPF used for fresh articles.
            bpf = next((b for b in bpfs if b.get("uniquename") == "newprocess"), None)
            if not bpf:
                # Fall back to any BPF that isn't translation/expired (those are for other lifecycles)
                bpf = next(
                    (
                        b
                        for b in bpfs
                        if "translation" not in (b.get("uniquename") or "")
                        and "expired" not in (b.get("uniquename") or "")
                    ),
                    None,
                )
            if not bpf:
                logger.info("No active Knowledge Article BPF found; skipping stage advancement.")
                self._kb_bpf = False
                return None

            bpf_id = bpf["workflowid"]
            unique = bpf["uniquename"]

            # Resolve the BPF instance entity set name via metadata (robust to non-standard pluralization)
            try:
                m = self._request(
                    "GET",
                    self._api(f"EntityDefinitions(LogicalName='{unique}')"),
                    params={"$select": "EntitySetName"},
                )
                m.raise_for_status()
                instance_set = m.json().get("EntitySetName") or f"{unique}s"
            except Exception:
                instance_set = f"{unique}s"

            r = self._request(
                "GET",
                self._api("processstages"),
                params={
                    "$filter": f"_processid_value eq {bpf_id}",
                    "$select": "processstageid,stagename",
                },
            )
            r.raise_for_status()
            stages = {s["stagename"]: s["processstageid"] for s in r.json().get("value", [])}

            logger.info(
                f"Resolved Knowledge Article BPF: {bpf['name']!r} "
                f"(instance set '{instance_set}') with stages: {list(stages.keys())}"
            )
            self._kb_bpf = (bpf_id, instance_set, stages)
            return self._kb_bpf
        except Exception as e:
            logger.warning(f"Could not resolve Knowledge Article BPF: {e}")
            self._kb_bpf = False
            return None

    def advance_bpf_to_publish_stage(self, article_id: str) -> None:
        """Best-effort: advance the article's BPF widget to the 'Publish' stage.

        Does NOT raise — any failure is logged at WARNING and swallowed, since
        BPF advancement is purely cosmetic; the article's state/status was
        already set by `publish_article`.
        """
        meta = self._get_kb_bpf_metadata()
        if not meta:
            return
        bpf_id, instance_set, stages = meta

        publish_stage_id = stages.get("Publish")
        if not publish_stage_id:
            logger.info("BPF has no 'Publish' stage; skipping stage advancement.")
            return

        traversed = ",".join(
            stages[name] for name in ("Author", "Review", "Publish") if name in stages
        )

        try:
            r = self._request(
                "GET",
                self._api(instance_set),
                params={
                    "$filter": (
                        f"_knowledgearticleid_value eq {article_id} and "
                        f"_processid_value eq {bpf_id}"
                    ),
                    "$select": "businessprocessflowinstanceid",
                    "$top": "1",
                },
            )
            r.raise_for_status()
            rows = r.json().get("value", [])

            if rows:
                instance_id = rows[0]["businessprocessflowinstanceid"]
                resp = self._request(
                    "PATCH",
                    self._api(f"{instance_set}({instance_id})"),
                    json={
                        "activestageid@odata.bind": f"/processstages({publish_stage_id})",
                        "traversedpath": traversed,
                    },
                )
            else:
                resp = self._request(
                    "POST",
                    self._api(instance_set),
                    json={
                        "knowledgearticleid@odata.bind": f"/knowledgearticles({article_id})",
                        "processid@odata.bind": f"/workflows({bpf_id})",
                        "activestageid@odata.bind": f"/processstages({publish_stage_id})",
                        "traversedpath": traversed,
                    },
                )

            if resp.status_code in (200, 201, 204):
                logger.info(f"Article {article_id}: BPF advanced to 'Publish' stage.")
            else:
                logger.warning(
                    f"Article {article_id}: BPF advancement returned "
                    f"{resp.status_code}: {resp.text[:300]}"
                )
        except Exception as e:
            logger.warning(f"Article {article_id}: BPF advancement failed: {e}")

    def get_article_counts_by_status(self) -> dict[str, int]:
        """Get a count of knowledge articles grouped by status.

        Returns a dict like: {"Draft": 5, "Approved": 2, "Published": 10, "Archived": 1}
        """
        status_labels = {
            0: "Draft",
            1: "Draft",
            2: "Draft",
            3: "Unapproved",
            4: "Unapproved",
            5: "Approved",
            6: "Scheduled",
            7: "Published",
            8: "Needs Review",
            9: "Updating",
            10: "Expired",
            11: "Rejected",
            12: "Archived",
            13: "Discarded",
        }

        # Query all articles grouped by statuscode
        url = self._api(
            "knowledgearticles?$apply=groupby((statuscode),aggregate($count as count))"
        )
        resp = self._request("GET", url)
        resp.raise_for_status()
        data = resp.json()

        counts: dict[str, int] = {}
        total = 0
        for item in data.get("value", []):
            code = item.get("statuscode", -1)
            count = item.get("count", 0)
            label = status_labels.get(code, f"Unknown ({code})")
            counts[label] = counts.get(label, 0) + count
            total += count

        counts["Total"] = total
        return counts
