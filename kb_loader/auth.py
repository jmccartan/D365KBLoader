"""Azure CLI authentication for Microsoft Graph and Dataverse APIs.

Uses `az account get-access-token` to acquire tokens — no app registration needed.
Automatically triggers `az login` if the user isn't authenticated.
"""

import json
import logging
import subprocess
import sys
import time

logger = logging.getLogger(__name__)

TOKEN_REFRESH_MARGIN_SECONDS = 300  # refresh 5 minutes before expiry


class AuthClient:
    """Manages authentication tokens via Azure CLI."""

    def __init__(self):
        self._tokens: dict[str, str] = {}
        self._token_expiry: dict[str, float] = {}
        self._ensure_az_cli()

    def _ensure_az_cli(self):
        """Verify Azure CLI is installed."""
        try:
            result = subprocess.run(
                ["az", "version", "--output", "json"],
                capture_output=True, text=True, timeout=15,
            )
            if result.returncode != 0:
                raise FileNotFoundError()
        except (FileNotFoundError, subprocess.TimeoutExpired):
            raise RuntimeError(
                "Azure CLI (az) is required but not found.\n"
                "Install it from: https://aka.ms/installazurecliwindows (Windows)\n"
                "                 https://aka.ms/installazurecli (Mac/Linux)\n"
                "Then run: az login"
            )

    def _login(self):
        """Run interactive az login."""
        logger.info("Opening browser for Azure login...")
        print("\n" + "=" * 60)
        print("  Azure login required. A browser window will open.")
        print("  Sign in with your Microsoft account.")
        print("=" * 60 + "\n")
        sys.stdout.flush()

        result = subprocess.run(
            ["az", "login", "--output", "none"],
            timeout=300,
        )
        if result.returncode != 0:
            raise RuntimeError("Azure login failed. Please run 'az login' manually.")

    def _is_token_valid(self, resource: str) -> bool:
        """Check if a cached token exists and hasn't expired."""
        if resource not in self._tokens:
            return False
        expiry = self._token_expiry.get(resource, 0)
        return time.time() < (expiry - TOKEN_REFRESH_MARGIN_SECONDS)

    def _get_token(self, resource: str) -> str:
        """Get an access token for the given resource via Azure CLI."""
        if self._is_token_valid(resource):
            return self._tokens[resource]

        # Try to get token; if it fails, trigger login and retry
        for attempt in range(2):
            result = subprocess.run(
                ["az", "account", "get-access-token", "--resource", resource, "--output", "json"],
                capture_output=True, text=True, timeout=30,
            )

            if result.returncode == 0:
                data = json.loads(result.stdout)
                token = data["accessToken"]
                self._tokens[resource] = token
                # Parse expiry — az cli returns "expiresOn" in local time format
                expires_on = data.get("expiresOn", "")
                if expires_on:
                    try:
                        from datetime import datetime
                        # Handle both formats: "2026-04-29 12:00:00.000000" and ISO 8601
                        expires_on = expires_on.replace("T", " ").split("+")[0].split("Z")[0]
                        dt = datetime.strptime(expires_on.strip(), "%Y-%m-%d %H:%M:%S.%f")
                        self._token_expiry[resource] = dt.timestamp()
                    except (ValueError, OSError):
                        # If parsing fails, set a conservative 30-minute expiry
                        self._token_expiry[resource] = time.time() + 1800
                else:
                    self._token_expiry[resource] = time.time() + 1800
                logger.info(f"Token acquired for {resource}")
                return token

            if attempt == 0:
                logger.info(f"Not logged in or token expired for {resource}. Triggering login...")
                self._login()
            else:
                raise RuntimeError(
                    f"Failed to get token for {resource}: {result.stderr.strip()}"
                )

        raise RuntimeError(f"Failed to get token for {resource}")

    def get_graph_token(self) -> str:
        """Get an access token for Microsoft Graph API."""
        return self._get_token("https://graph.microsoft.com")

    def get_dataverse_token(self, dataverse_url: str) -> str:
        """Get an access token for Dataverse Web API."""
        resource = dataverse_url.rstrip("/")
        return self._get_token(resource)
