"""SharePoint client using Microsoft Graph API.

Resolves SharePoint folder URLs to drive items and recursively enumerates Word files.
"""

import logging
import time
from dataclasses import dataclass
from pathlib import Path
from urllib.parse import urlparse, unquote
import requests
from kb_loader.auth import AuthClient
from kb_loader.converter import SUPPORTED_EXTENSIONS

logger = logging.getLogger(__name__)

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
MAX_RETRIES = 3
RETRY_BASE_DELAY = 2


@dataclass
class SharePointFile:
    """Represents a Word file (.docx or .doc) found in SharePoint."""
    name: str
    item_id: str
    drive_id: str
    download_url: str
    relative_path: str  # path relative to the target folder
    last_modified: str
    size: int


class SharePointClient:
    """Client for enumerating and downloading files from SharePoint via Graph API."""

    def __init__(self, auth: AuthClient):
        self.auth = auth
        self.session = requests.Session()

    def _headers(self) -> dict:
        return {
            "Authorization": f"Bearer {self.auth.get_graph_token()}",
            "Accept": "application/json",
        }

    def _get(self, url: str, params: dict | None = None) -> dict:
        """GET with retry and throttle handling."""
        for attempt in range(MAX_RETRIES):
            resp = self.session.get(url, headers=self._headers(), params=params)
            if resp.status_code == 429:
                retry_after = int(resp.headers.get("Retry-After", RETRY_BASE_DELAY * (attempt + 1)))
                logger.warning(f"Throttled by Graph API. Waiting {retry_after}s...")
                time.sleep(retry_after)
                continue
            if resp.status_code >= 500:
                delay = RETRY_BASE_DELAY * (attempt + 1)
                logger.warning(f"Server error {resp.status_code}. Retrying in {delay}s...")
                time.sleep(delay)
                continue
            resp.raise_for_status()
            return resp.json()
        raise RuntimeError(f"Graph API request failed after {MAX_RETRIES} retries: {url}")

    def _parse_sharepoint_url(self, folder_url: str) -> tuple[str, str, str]:
        """Parse a SharePoint folder URL into (hostname, site_path, folder_path).

        Supports URLs like:
          https://tenant.sharepoint.com/sites/MySite/Shared Documents/MyFolder/Sub
          https://tenant.sharepoint.com/sites/MySite/Documents/MyFolder
        """
        parsed = urlparse(folder_url)
        hostname = parsed.hostname
        path_parts = unquote(parsed.path).strip("/").split("/")

        # Find the site path (e.g., "sites/MySite" or "teams/MyTeam")
        site_path = None
        doc_lib_start = None
        for i, part in enumerate(path_parts):
            if part.lower() in ("sites", "teams") and i + 1 < len(path_parts):
                site_path = f"{part}/{path_parts[i + 1]}"
                doc_lib_start = i + 2
                break

        if not site_path:
            raise ValueError(
                f"Cannot determine site path from URL: {folder_url}\n"
                "Expected format: https://tenant.sharepoint.com/sites/SiteName/Library/Folder"
            )

        # Everything after the site path: first segment is the library, rest is folder path
        remaining = path_parts[doc_lib_start:]
        if not remaining:
            raise ValueError(
                f"Cannot determine document library from URL: {folder_url}\n"
                "Expected at least a document library name after the site path."
            )

        # The library name is the first remaining segment; folder_path is the rest
        library_name = remaining[0]
        folder_path = "/".join(remaining[1:]) if len(remaining) > 1 else ""

        return hostname, site_path, library_name, folder_path

    def _resolve_site_id(self, hostname: str, site_path: str) -> str:
        """Resolve a SharePoint site to its Graph site ID."""
        url = f"{GRAPH_BASE}/sites/{hostname}:/{site_path}"
        data = self._get(url)
        site_id = data["id"]
        logger.info(f"Resolved site ID: {site_id}")
        return site_id

    def _resolve_drive(self, site_id: str, library_name: str) -> str:
        """Find the drive ID for a document library by name."""
        url = f"{GRAPH_BASE}/sites/{site_id}/drives"
        data = self._get(url)

        for drive in data.get("value", []):
            if drive["name"].lower() == library_name.lower():
                logger.info(f"Resolved drive '{library_name}' → {drive['id']}")
                return drive["id"]

        # Fallback: try matching on webUrl containing the library name
        for drive in data.get("value", []):
            if library_name.lower().replace(" ", "%20") in drive.get("webUrl", "").lower():
                logger.info(f"Resolved drive via webUrl match '{library_name}' → {drive['id']}")
                return drive["id"]

        available = [d["name"] for d in data.get("value", [])]
        raise ValueError(
            f"Document library '{library_name}' not found. Available libraries: {available}"
        )

    def _resolve_folder_item(self, drive_id: str, folder_path: str) -> str | None:
        """Resolve a folder path within a drive to a driveItem ID. Returns None for root."""
        if not folder_path:
            return None  # root of the drive
        url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{folder_path}"
        data = self._get(url)
        item_id = data["id"]
        logger.info(f"Resolved folder path '{folder_path}' → {item_id}")
        return item_id

    def _list_children(self, drive_id: str, item_id: str | None) -> list[dict]:
        """List all children of a drive item, handling pagination."""
        if item_id:
            url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/children"
        else:
            url = f"{GRAPH_BASE}/drives/{drive_id}/root/children"

        items = []
        while url:
            data = self._get(url)
            items.extend(data.get("value", []))
            url = data.get("@odata.nextLink")
        return items

    def enumerate_docx_files(self, folder_url: str) -> list[SharePointFile]:
        """Recursively find all Word files (.docx, .doc) in the given SharePoint folder."""
        hostname, site_path, library_name, folder_path = self._parse_sharepoint_url(folder_url)

        logger.info(f"Resolving SharePoint site: {hostname}/{site_path}")
        site_id = self._resolve_site_id(hostname, site_path)
        drive_id = self._resolve_drive(site_id, library_name)
        root_item_id = self._resolve_folder_item(drive_id, folder_path)

        files = []
        self._recurse_folder(drive_id, root_item_id, "", files)

        logger.info(f"Found {len(files)} Word file(s) in SharePoint.")
        return files

    def _recurse_folder(
        self, drive_id: str, item_id: str | None, relative_path: str, results: list[SharePointFile]
    ):
        """Recursively enumerate Word files in a folder."""
        children = self._list_children(drive_id, item_id)

        for child in children:
            child_name = child["name"]

            if "folder" in child:
                # Recurse into subfolder
                sub_path = f"{relative_path}/{child_name}" if relative_path else child_name
                logger.debug(f"Entering subfolder: {sub_path}")
                self._recurse_folder(drive_id, child["id"], sub_path, results)

            elif "file" in child and Path(child_name).suffix.lower() in SUPPORTED_EXTENSIONS:
                download_url = child.get("@microsoft.graph.downloadUrl", "")
                results.append(
                    SharePointFile(
                        name=child_name,
                        item_id=child["id"],
                        drive_id=drive_id,
                        download_url=download_url,
                        relative_path=relative_path,
                        last_modified=child.get("lastModifiedDateTime", ""),
                        size=child.get("size", 0),
                    )
                )
                logger.debug(f"Found: {relative_path}/{child_name}")

    def download_file(self, file: SharePointFile) -> bytes:
        """Download a SharePoint file's content as bytes."""
        if file.download_url:
            resp = self.session.get(file.download_url)
        else:
            url = f"{GRAPH_BASE}/drives/{file.drive_id}/items/{file.item_id}/content"
            resp = self.session.get(url, headers=self._headers())

        resp.raise_for_status()
        return resp.content
