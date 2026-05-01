"""Authentication for Microsoft Graph and Dataverse APIs.

Two methods, transparent to callers:
  - MSAL interactive browser login — used when AZURE_CLIENT_ID is configured.
    Uses the "organizations" authority so users never need to specify a tenant.
  - Azure CLI `az account get-access-token` — fallback when MSAL is not configured.
    Auto-detects the tenant from the existing az session.

Public API:
  AuthClient(client_id="", tenant_id="")  — both args optional
  .get_graph_token()
  .get_dataverse_token(dataverse_url)
  .get_signed_in_user() -> str | None
  .sign_out()
  .set_device_code_callback(cb) — when set, az login uses device-code flow and
    invokes cb(code, url) so the UI can display the code prominently. Otherwise
    az login uses its default interactive browser flow.
"""

import json
import logging
import os
import platform
import re
import subprocess
import sys
import time
from pathlib import Path
from typing import Callable, Optional

logger = logging.getLogger(__name__)

TOKEN_REFRESH_MARGIN_SECONDS = 300
_WINDOWS = platform.system() == "Windows"
_CACHE_DIR = Path.home() / ".d365kbloader"
_TOKEN_CACHE_FILE = _CACHE_DIR / "msal_token_cache.json"

# Graph scopes needed for SharePoint file enumeration and download
_GRAPH_SCOPES = [
    "https://graph.microsoft.com/Sites.Read.All",
    "https://graph.microsoft.com/Files.Read.All",
]

# DeviceCodeCallback receives (code, sign_in_url)
DeviceCodeCallback = Callable[[str, str], None]


class AuthClient:
    """Manages authentication tokens via MSAL or Azure CLI."""

    def __init__(self, client_id: str = "", tenant_id: str = ""):
        self._tokens: dict[str, str] = {}
        self._token_expiry: dict[str, float] = {}
        self._msal_app = None
        self._msal_cache = None
        self._device_code_callback: Optional[DeviceCodeCallback] = None

        # Allow constructor args; fall back to environment variables for back-compat
        self._client_id = client_id or os.environ.get("AZURE_CLIENT_ID", "")
        self._tenant_id = tenant_id or os.environ.get("AZURE_TENANT_ID", "")

        if self._client_id:
            self._method = "msal"
            self._init_msal()
        else:
            self._method = "az_cli"
            self._ensure_az_cli()
            if not self._tenant_id:
                self._tenant_id = self._detect_tenant_az()
                if self._tenant_id:
                    logger.info(f"Auto-detected tenant: {self._tenant_id}")

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

        # Use "organizations" authority when no tenant is specified — works
        # for any work/school account without forcing the user to pick a tenant.
        authority = (
            f"https://login.microsoftonline.com/{self._tenant_id}"
            if self._tenant_id
            else "https://login.microsoftonline.com/organizations"
        )

        self._msal_cache = msal.SerializableTokenCache()
        if _TOKEN_CACHE_FILE.exists():
            try:
                self._msal_cache.deserialize(_TOKEN_CACHE_FILE.read_text())
            except Exception:
                logger.warning("Could not load MSAL cache; starting fresh.")

        try:
            self._msal_app = msal.PublicClientApplication(
                client_id=self._client_id,
                authority=authority,
                token_cache=self._msal_cache,
            )
        except ValueError as e:
            raise RuntimeError(
                f"MSAL initialization failed.\n"
                f"  Client ID: {self._client_id}\n"
                f"  Authority: {authority}\n"
                f"  Detail: {e}"
            )
        logger.info(f"MSAL initialized (authority: {authority})")

    def _save_msal_cache(self):
        """Persist MSAL token cache to disk."""
        if self._msal_cache and self._msal_cache.has_state_changed:
            _CACHE_DIR.mkdir(parents=True, exist_ok=True)
            _TOKEN_CACHE_FILE.write_text(self._msal_cache.serialize())

    def _get_token_msal(self, scopes: list[str]) -> str:
        """Acquire token via MSAL — silent first, then interactive browser."""
        accounts = self._msal_app.get_accounts()
        result = None

        if accounts:
            result = self._msal_app.acquire_token_silent(scopes, account=accounts[0])

        if not result or "access_token" not in result:
            logger.info("Opening browser for sign-in...")
            result = self._msal_app.acquire_token_interactive(scopes=scopes)

        self._save_msal_cache()

        if "access_token" in result:
            logger.info(f"Token acquired via MSAL")
            return result["access_token"]

        error = result.get("error_description", result.get("error", "Unknown error"))
        raise RuntimeError(f"Sign-in failed: {error}")

    # ── Azure CLI auth ─────────────────────────────────────────────────

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
                "Azure CLI is not installed.\n"
                "Install from: https://aka.ms/installazurecli\n"
                "  Windows: winget install Microsoft.AzureCLI\n"
                "  Mac:     brew install azure-cli\n\n"
                "Alternatively, configure MSAL by setting AZURE_CLIENT_ID."
            )

    def _detect_tenant_az(self) -> str:
        """Try to detect the tenant ID from the current az CLI session."""
        try:
            result = subprocess.run(
                ["az", "account", "show", "--output", "json"],
                capture_output=True, text=True, timeout=15,
                shell=_WINDOWS,
            )
            if result.returncode == 0:
                data = json.loads(result.stdout)
                return data.get("tenantId", "")
        except (json.JSONDecodeError, subprocess.TimeoutExpired, FileNotFoundError):
            pass
        return ""

    def _detect_user_az(self) -> str:
        """Get the signed-in user from the az CLI session."""
        try:
            result = subprocess.run(
                ["az", "account", "show", "--output", "json"],
                capture_output=True, text=True, timeout=15,
                shell=_WINDOWS,
            )
            if result.returncode == 0:
                data = json.loads(result.stdout)
                user = data.get("user", {})
                return user.get("name", "") if isinstance(user, dict) else ""
        except (json.JSONDecodeError, subprocess.TimeoutExpired, FileNotFoundError):
            pass
        return ""

    def _extract_tenant_from_error(self, stderr: str) -> str:
        """Parse tenant ID from az CLI error messages."""
        match = re.search(r'--tenant\s+"([0-9a-f-]{36})"', stderr)
        return match.group(1) if match else ""

    def _login_az(self, resource: str | None = None):
        """Run az login. Uses device-code flow if a callback is registered, else
        falls back to az's default interactive browser flow."""
        if self._device_code_callback is not None:
            self._login_az_device_code(resource)
        else:
            self._login_az_interactive(resource)

    def _login_az_interactive(self, resource: str | None = None):
        """Run az's default interactive browser login."""
        cmd = ["az", "login", "--allow-no-subscriptions"]
        if self._tenant_id:
            cmd.extend(["--tenant", self._tenant_id])
        if resource:
            cmd.extend(["--scope", f"{resource.rstrip('/')}/.default"])
        cmd.extend(["--output", "none"])

        result = subprocess.run(cmd, timeout=300, shell=_WINDOWS)
        if result.returncode != 0:
            raise RuntimeError(
                "Sign-in was cancelled or failed. "
                "Try clicking Sign In again."
            )

    def _login_az_device_code(self, resource: str | None = None):
        """Run az login with --use-device-code, parse the code from stdout, and
        invoke the registered callback so the UI can display it prominently."""
        cmd = ["az", "login", "--use-device-code", "--allow-no-subscriptions"]
        if self._tenant_id:
            cmd.extend(["--tenant", self._tenant_id])
        if resource:
            cmd.extend(["--scope", f"{resource.rstrip('/')}/.default"])
        cmd.extend(["--output", "none"])

        proc = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            shell=_WINDOWS,
            bufsize=1,
        )

        # Match az output like:
        #   "To sign in, use a web browser to open the page https://microsoft.com/devicelogin
        #    and enter the code ABC123XYZ to authenticate."
        code_pattern = re.compile(
            r"open the page\s+(\S+)\s+and enter the code\s+(\S+)",
            re.IGNORECASE,
        )
        code_emitted = False
        captured_lines: list[str] = []
        start = time.time()
        TIMEOUT = 600  # 10 minutes for the user to enter the code

        try:
            while True:
                if proc.stdout is None:
                    break
                line = proc.stdout.readline()
                if line:
                    captured_lines.append(line)
                    if not code_emitted:
                        m = code_pattern.search(line)
                        if m:
                            url, code = m.group(1), m.group(2)
                            try:
                                self._device_code_callback(code, url)
                            except Exception as cb_err:
                                logger.warning(f"Device code callback error: {cb_err}")
                            code_emitted = True
                if proc.poll() is not None:
                    # process exited — drain any remaining output
                    if proc.stdout is not None:
                        rest = proc.stdout.read()
                        if rest:
                            captured_lines.append(rest)
                    break
                if time.time() - start > TIMEOUT:
                    proc.kill()
                    raise RuntimeError(
                        "Sign-in timed out (10 minutes). "
                        "Click Sign In again to retry."
                    )
        finally:
            try:
                proc.wait(timeout=5)
            except subprocess.TimeoutExpired:
                proc.kill()

        if proc.returncode != 0:
            output = "".join(captured_lines).strip()
            raise RuntimeError(
                "Sign-in failed. Click Sign In again to retry.\n"
                f"Detail: {output[-500:]}" if output else "Sign-in failed. Click Sign In again to retry."
            )

    def _get_token_az(self, resource: str) -> str:
        """Get an access token for the given resource via Azure CLI."""
        cache_key = f"az:{resource}"
        if cache_key in self._tokens:
            expiry = self._token_expiry.get(cache_key, 0)
            if time.time() < (expiry - TOKEN_REFRESH_MARGIN_SECONDS):
                return self._tokens[cache_key]

        for attempt in range(2):
            cmd = [
                "az", "account", "get-access-token",
                "--resource", resource, "--output", "json",
            ]
            if self._tenant_id:
                cmd.extend(["--tenant", self._tenant_id])

            result = subprocess.run(
                cmd, capture_output=True, text=True, timeout=30,
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
                if not self._tenant_id:
                    detected = self._extract_tenant_from_error(result.stderr)
                    if detected:
                        self._tenant_id = detected
                        logger.info(f"Auto-detected tenant from error: {self._tenant_id}")
                logger.info(f"Triggering sign-in for {resource}...")
                self._login_az(resource)
            else:
                raise RuntimeError(
                    "Could not get a token from Azure CLI.\n"
                    "Try signing out and signing in again.\n"
                    f"Detail: {result.stderr.strip()}"
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

    def get_sharepoint_token(self, hostname: str) -> str:
        """Get an access token for the SharePoint REST API on a specific tenant host.

        Used to resolve sharing links by following SharePoint's redirect, when
        Microsoft Graph's /shares endpoint can't be used with the available scope.
        """
        resource = f"https://{hostname.rstrip('/')}"
        if self._method == "msal":
            return self._get_token_msal([f"{resource}/AllSites.Read"])
        return self._get_token_az(resource)

    def get_signed_in_user(self) -> str | None:
        """Return the username/email of the signed-in user, or None if not signed in."""
        if self._method == "msal":
            if not self._msal_app:
                return None
            accounts = self._msal_app.get_accounts()
            if accounts:
                return accounts[0].get("username")
            return None
        # az CLI
        return self._detect_user_az() or None

    def sign_out(self):
        """Sign out — clears MSAL cache. (Does NOT log out of az CLI globally.)"""
        if self._method == "msal" and self._msal_app:
            for account in self._msal_app.get_accounts():
                self._msal_app.remove_account(account)
            if _TOKEN_CACHE_FILE.exists():
                try:
                    _TOKEN_CACHE_FILE.unlink()
                except OSError:
                    pass
            self._tokens.clear()
            self._token_expiry.clear()
            logger.info("Signed out (MSAL cache cleared).")
        else:
            # For az CLI, we don't touch global state. Just clear our in-memory cache.
            self._tokens.clear()
            self._token_expiry.clear()
            logger.info(
                "In-memory tokens cleared. "
                "To fully sign out of Azure CLI, run 'az logout' manually."
            )

    @property
    def method(self) -> str:
        """Current auth method: 'msal' or 'az_cli'."""
        return self._method

    def set_device_code_callback(self, cb: Optional[DeviceCodeCallback]):
        """Register a callback to receive device-code sign-in info.

        When set (non-None), az CLI sign-in uses --use-device-code flow:
        the callback is invoked with (code, sign_in_url) so the UI can
        prominently display the code instead of relying on a browser
        window that might be hidden behind other apps.

        Set to None to revert to az's default interactive browser flow.
        """
        self._device_code_callback = cb

