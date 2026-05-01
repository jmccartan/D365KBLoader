# =====================================================================
# D365 Knowledge Base Loader
# Copyright (c) 2026. All rights reserved.
# Licensed under the MIT License. See the LICENSE file in the project
# root for the full text.
# =====================================================================

"""Configuration management from environment variables and .env file."""

import os
from pathlib import Path
from urllib.parse import urlparse
from dataclasses import dataclass, field
from dotenv import load_dotenv


@dataclass
class Config:
    dataverse_url: str
    output_dir: str
    existing_article_mode: str

    # Input source - one of these will be set
    sharepoint_folder_url: str = ""
    local_folder: str = ""

    @property
    def dataverse_api_url(self) -> str:
        return f"{self.dataverse_url.rstrip('/')}/api/data/v9.2"

    @property
    def input_mode(self) -> str:
        return "local" if self.local_folder else "sharepoint"


def load_config(
    sharepoint_url: str | None = None,
    local_folder: str | None = None,
    output_dir: str | None = None,
    existing_mode: str | None = None,
) -> Config:
    """Load configuration from .env file and optional CLI overrides."""
    env_path = Path(".env")
    if not env_path.exists():
        example_path = Path(".env.example")
        hint = " Copy .env.example to .env and configure your settings." if example_path.exists() else ""
        raise ValueError(f"No .env file found in {Path.cwd()}.{hint}")
    load_dotenv(dotenv_path=env_path)

    dataverse_url = os.getenv("DATAVERSE_URL", "")
    sp_url = sharepoint_url or os.getenv("SHAREPOINT_FOLDER_URL", "")
    local = local_folder or os.getenv("LOCAL_FOLDER", "")
    out = output_dir or os.getenv("OUTPUT_DIR", "./output")
    existing = existing_mode or os.getenv("EXISTING_ARTICLE_MODE", "skip")

    if not dataverse_url:
        raise ValueError("DATAVERSE_URL is required. Set it in .env or environment.")

    parsed = urlparse(dataverse_url)
    if parsed.scheme != "https" or not parsed.hostname:
        raise ValueError(
            f"DATAVERSE_URL must be a valid HTTPS URL (e.g. https://your-org.crm.dynamics.com), "
            f"got: {dataverse_url}"
        )
    if not sp_url and not local:
        raise ValueError(
            "An input source is required. Use --sharepoint-url or --local-folder."
        )
    if sp_url and local:
        raise ValueError("Specify either --sharepoint-url or --local-folder, not both.")
    if existing not in ("skip", "update", "duplicate"):
        raise ValueError(
            f"EXISTING_ARTICLE_MODE must be 'skip', 'update', or 'duplicate', got '{existing}'"
        )

    return Config(
        dataverse_url=dataverse_url,
        sharepoint_folder_url=sp_url,
        local_folder=local,
        output_dir=out,
        existing_article_mode=existing,
    )
