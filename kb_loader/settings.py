# =====================================================================
# D365 Knowledge Base Loader
# Copyright (c) 2026 John McCartan
# Licensed under the MIT License. See the LICENSE file in the project
# root for the full text.
# =====================================================================

"""User-facing settings stored in the user profile directory.

Settings live in `~/.d365kbloader/settings.json` so they persist across the
project folder being moved or re-cloned. The legacy `.env` file in the project
folder is read as a fallback for backward compatibility.

Settings are NOT secrets — they hold URLs, paths, and choices. Auth tokens
are handled separately in `auth.py`.
"""

import json
import logging
import os
from dataclasses import dataclass, asdict, field
from pathlib import Path

from dotenv import dotenv_values

logger = logging.getLogger(__name__)

SETTINGS_DIR = Path.home() / ".d365kbloader"
SETTINGS_FILE = SETTINGS_DIR / "settings.json"


@dataclass
class Settings:
    """Persistent user settings."""
    dataverse_url: str = ""
    sharepoint_folder_url: str = ""
    local_folder: str = ""
    output_dir: str = "./output"
    existing_article_mode: str = "skip"  # skip | update | duplicate

    # Auth — only set if MSAL custom app registration is in use.
    # Empty values mean "use Azure CLI" (no tenant config needed).
    azure_client_id: str = ""
    azure_tenant_id: str = ""

    @property
    def input_mode(self) -> str:
        return "local" if self.local_folder else "sharepoint"

    def is_valid_for_run(self) -> tuple[bool, str]:
        """Check if settings are sufficient to run a load. Returns (ok, error_msg)."""
        if not self.dataverse_url.strip():
            return False, "Dataverse URL is required."
        if not self.dataverse_url.startswith("https://"):
            return False, "Dataverse URL must start with https://"
        if not self.sharepoint_folder_url.strip() and not self.local_folder.strip():
            return False, "Pick a SharePoint folder URL or a local folder."
        if self.sharepoint_folder_url.strip() and self.local_folder.strip():
            return False, "Use either a SharePoint URL or a local folder, not both."
        if self.existing_article_mode not in ("skip", "update", "duplicate"):
            return False, f"Existing-article mode must be skip/update/duplicate."
        if self.local_folder.strip() and not Path(self.local_folder).is_dir():
            return False, f"Local folder not found: {self.local_folder}"
        return True, ""


def load_settings() -> Settings:
    """Load settings from the user profile JSON, then merge in .env overrides.

    Order of precedence (last wins):
      1. Defaults
      2. ~/.d365kbloader/settings.json
      3. ./.env  (for back-compat with the old CLI workflow)
      4. Process environment variables
    """
    s = Settings()

    if SETTINGS_FILE.exists():
        try:
            data = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
            for k, v in data.items():
                if hasattr(s, k) and isinstance(v, str):
                    setattr(s, k, v)
        except (json.JSONDecodeError, OSError) as e:
            logger.warning(f"Could not read {SETTINGS_FILE}: {e}")

    # .env overlay (only if file exists in the cwd)
    env_file = Path(".env")
    env_values = {}
    if env_file.exists():
        try:
            env_values = dotenv_values(env_file)
        except Exception as e:
            logger.warning(f"Could not read .env: {e}")

    def _pick(*keys: str) -> str:
        for k in keys:
            v = os.environ.get(k) or env_values.get(k)
            if v:
                return v.strip()
        return ""

    if v := _pick("DATAVERSE_URL"): s.dataverse_url = v
    if v := _pick("SHAREPOINT_FOLDER_URL"): s.sharepoint_folder_url = v
    if v := _pick("LOCAL_FOLDER"): s.local_folder = v
    if v := _pick("OUTPUT_DIR"): s.output_dir = v
    if v := _pick("EXISTING_ARTICLE_MODE"): s.existing_article_mode = v
    if v := _pick("AZURE_CLIENT_ID"): s.azure_client_id = v
    if v := _pick("AZURE_TENANT_ID"): s.azure_tenant_id = v

    return s


def save_settings(s: Settings) -> Path:
    """Write settings to the user profile JSON."""
    SETTINGS_DIR.mkdir(parents=True, exist_ok=True)
    payload = asdict(s)
    SETTINGS_FILE.write_text(
        json.dumps(payload, indent=2, sort_keys=True),
        encoding="utf-8",
    )
    logger.info(f"Settings saved to {SETTINGS_FILE}")
    return SETTINGS_FILE
