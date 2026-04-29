"""Azure CLI authentication for Microsoft Graph and Dataverse APIs.

Uses `az account get-access-token` to acquire tokens — no app registration needed.
Automatically triggers `az login` if the user isn't authenticated.
"""

import json
import logging
import subprocess
import sys

logger = logging.getLogger(__name__)


class AuthClient:
    """Manages authentication tokens via Azure CLI."""

    def __init__(self):
        self._tokens: dict[str, str] = {}
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

    def _get_token(self, resource: str) -> str:
        """Get an access token for the given resource via Azure CLI."""
        if resource in self._tokens:
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
