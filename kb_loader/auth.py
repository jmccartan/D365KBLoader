"""Authentication for Microsoft Graph and Dataverse APIs.

Supports two methods (controlled by AUTH_METHOD env var):
  - "msal"   : MSAL interactive browser login (requires AZURE_CLIENT_ID + AZURE_TENANT_ID)
  - "az_cli" : Azure CLI `az account get-access-token` (requires az login + subscription)
  - "auto"   : (default) try MSAL if configured, fall back to az CLI
"""

import json
import logging
import os
import platform
import subprocess
import sys
import time
from pathlib import Path

logger = logging.getLogger(__name__)

TOKEN_REFRESH_MARGIN_SECONDS = 300  # refresh 5 minutes before expiry
_WINDOWS = platform.system() == "Windows"
_CACHE_DIR = Path.home() / ".d365kbloader"
_TOKEN_CACHE_FILE = _CACHE_DIR / "msal_token_cache.json"

# Graph scopes needed for SharePoint file enumeration and download
_GRAPH_SCOPES = [
    "https://graph.microsoft.com/Sites.Read.All",
    "https://graph.microsoft.com/Files.Read.All",
]


class AuthClient:
    """Manages authentication tokens via MSAL or Azure CLI."""

    def __init__(self):
        self._tokens: dict[str, str] = {}
        self._token_expiry: dict[str, float] = {}
        self._msal_app = None
        self._msal_cache = None

        method = os.environ.get("AUTH_METHOD", "auto").lower()
        if method not in ("auto", "msal", "az_cli"):
            raise ValueError(f"AUTH_METHOD must be 'auto', 'msal', or 'az_cli', got '{method}'")

        self._method = method
        self._client_id = os.environ.get("AZURE_CLIENT_ID", "")
        self._tenant_id = os.environ.get("AZURE_TENANT_ID", "")

        if method == "msal" and not self._client_id:
            raise RuntimeError(
                "AUTH_METHOD=msal requires AZURE_CLIENT_ID.\n"
                "Register a public client app in Entra ID and set AZURE_CLIENT_ID in .env"
            )

        # Resolve effective method
        if method == "auto":
            if self._client_id and self._tenant_id:
                self._method = "msal"
                logger.info("Auth: using MSAL interactive (AZURE_CLIENT_ID configured)")
            else:
                self._method = "az_cli"
                logger.info("Auth: using Azure CLI (set AZURE_CLIENT_ID + AZURE_TENANT_ID for MSAL)")

        if self._method == "msal":
            self._init_msal()
        else:
            self._ensure_az_cli()

    # ── MSAL interactive auth ──────────────────────────────────────────

    def _init_msal(self):
        """Initialize MSAL public client app with persistent token cache."""
        try:
            import msal
        except ImportError:
            raise RuntimeError(
                "MSAL library is required for interactive auth.\n"
                "Install it: pip install msal"
            )

        authority = f"https://login.microsoftonline.com/{self._tenant_id}"

        self._msal_cache = msal.SerializableTokenCache()
        if _TOKEN_CACHE_FILE.exists():
            self._msal_cache.deserialize(_TOKEN_CACHE_FILE.read_text())

        try:
            self._msal_app = msal.PublicClientApplication(
                client_id=self._client_id,
                authority=authority,
                token_cache=self._msal_cache,
            )
        except ValueError as e:
            raise RuntimeError(
                f"MSAL initialization failed — check AZURE_TENANT_ID.\n"
                f"Current value: {self._tenant_id}\n"
                f"Detail: {e}"
            )
        logger.info(f"MSAL initialized (tenant: {self._tenant_id})")

    def _save_msal_cache(self):
        """Persist MSAL token cache to disk."""
        if self._msal_cache and self._msal_cache.has_state_changed:
            _CACHE_DIR.mkdir(parents=True, exist_ok=True)
            _TOKEN_CACHE_FILE.write_text(self._msal_cache.serialize())

    def _get_token_msal(self, scopes: list[str]) -> str:
        """Acquire token via MSAL — silent first, then interactive."""
        accounts = self._msal_app.get_accounts()
        result = None

        if accounts:
            result = self._msal_app.acquire_token_silent(scopes, account=accounts[0])

        if not result or "access_token" not in result:
            logger.info("Opening browser for Azure login...")
            print("\n" + "=" * 60)
            print("  Azure login required. A browser window will open.")
            print("  Sign in with your Microsoft account.")
            print("=" * 60 + "\n")
            sys.stdout.flush()
            result = self._msal_app.acquire_token_interactive(scopes=scopes)

        self._save_msal_cache()

        if "access_token" in result:
            logger.info(f"Token acquired via MSAL for scopes: {scopes}")
            return result["access_token"]

        error = result.get("error_description", result.get("error", "Unknown error"))
        raise RuntimeError(f"MSAL authentication failed: {error}")

    # ── Azure CLI auth (fallback) ──────────────────────────────────────

    def _ensure_az_cli(self):
        """Verify Azure CLI is installed."""
        try:
            result = subprocess.run(
                ["az", "version", "--output", "json"],
                capture_output=True, text=True, timeout=15,
                shell=_WINDOWS,
            )
            if result.returncode != 0:
                raise FileNotFoundError()
        except (FileNotFoundError, subprocess.TimeoutExpired):
            raise RuntimeError(
                "Azure CLI (az) is required but not found.\n"
                "Install it from: https://aka.ms/installazurecliwindows (Windows)\n"
                "                 https://aka.ms/installazurecli (Mac/Linux)\n"
                "Then run: az login\n\n"
                "Alternatively, set AZURE_CLIENT_ID and AZURE_TENANT_ID in .env to use MSAL instead."
            )

    def _login_az(self):
        """Run interactive az login."""
        logger.info("Opening browser for Azure login...")
        print("\n" + "=" * 60)
        print("  Azure login required. A browser window will open.")
        print("  Sign in with your Microsoft account.")
        print("=" * 60 + "\n")
        sys.stdout.flush()

        result = subprocess.run(
            ["az", "login", "--allow-no-subscriptions", "--output", "none"],
            timeout=300,
            shell=_WINDOWS,
        )
        if result.returncode != 0:
            raise RuntimeError("Azure login failed. Please run 'az login' manually.")

    def _get_token_az(self, resource: str) -> str:
        """Get an access token for the given resource via Azure CLI."""
        cache_key = f"az:{resource}"
        if cache_key in self._tokens:
            expiry = self._token_expiry.get(cache_key, 0)
            if time.time() < (expiry - TOKEN_REFRESH_MARGIN_SECONDS):
                return self._tokens[cache_key]

        for attempt in range(2):
            result = subprocess.run(
                ["az", "account", "get-access-token", "--resource", resource, "--output", "json"],
                capture_output=True, text=True, timeout=30,
                shell=_WINDOWS,
            )

            if result.returncode == 0:
                data = json.loads(result.stdout)
                token = data["accessToken"]
                self._tokens[cache_key] = token
                expires_on = data.get("expiresOn", "")
                if expires_on:
                    try:
                        from datetime import datetime
                        expires_on = expires_on.replace("T", " ").split("+")[0].split("Z")[0]
                        dt = datetime.strptime(expires_on.strip(), "%Y-%m-%d %H:%M:%S.%f")
                        self._token_expiry[cache_key] = dt.timestamp()
                    except (ValueError, OSError):
                        self._token_expiry[cache_key] = time.time() + 1800
                else:
                    self._token_expiry[cache_key] = time.time() + 1800
                logger.info(f"Token acquired via az CLI for {resource}")
                return token

            if attempt == 0:
                logger.info(f"Not logged in or token expired for {resource}. Triggering login...")
                self._login_az()
            else:
                raise RuntimeError(
                    f"Failed to get token for {resource}: {result.stderr.strip()}"
                )

        raise RuntimeError(f"Failed to get token for {resource}")

    # ── Public API ─────────────────────────────────────────────────────

    def get_graph_token(self) -> str:
        """Get an access token for Microsoft Graph API."""
        if self._method == "msal":
            return self._get_token_msal(_GRAPH_SCOPES)
        return self._get_token_az("https://graph.microsoft.com")

    def get_dataverse_token(self, dataverse_url: str) -> str:
        """Get an access token for Dataverse Web API."""
        resource = dataverse_url.rstrip("/")
        if self._method == "msal":
            return self._get_token_msal([f"{resource}/user_impersonation"])
        return self._get_token_az(resource)
