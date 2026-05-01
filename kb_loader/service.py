# =====================================================================
# D365 Knowledge Base Loader
# Copyright (c) 2026. All rights reserved.
# Licensed under the MIT License. See the LICENSE file in the project
# root for the full text.
# =====================================================================

"""Service layer — the core load operation, callable from CLI or GUI.

The service emits structured progress events via a callback so the caller
(CLI or GUI) can render them however they like. It never calls print() or
sys.exit() directly.
"""

import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Optional

from kb_loader.auth import AuthClient
from kb_loader.converter import convert_to_html, save_html_file, SUPPORTED_EXTENSIONS
from kb_loader.dataverse_client import DataverseClient
from kb_loader.run_log import RunLog
from kb_loader.settings import Settings

logger = logging.getLogger(__name__)


@dataclass
class DocxFile:
    """A Word document (.docx or .doc) to process, from either local or SharePoint source."""
    name: str
    relative_path: str  # subfolder path relative to root
    source_display: str  # for logging

    # One of these will be set
    local_path: Optional[Path] = None
    sp_file: object = None  # SharePointFile


@dataclass
class LoadConfig:
    """Resolved configuration for a single load operation."""
    dataverse_url: str
    sharepoint_folder_url: str = ""
    local_folder: str = ""
    output_dir: str = "./output"
    existing_article_mode: str = "skip"
    dry_run: bool = False

    @property
    def input_mode(self) -> str:
        return "local" if self.local_folder else "sharepoint"

    @property
    def dataverse_api_url(self) -> str:
        return f"{self.dataverse_url.rstrip('/')}/api/data/v9.2"

    @classmethod
    def from_settings(cls, s: Settings, dry_run: bool = False) -> "LoadConfig":
        return cls(
            dataverse_url=s.dataverse_url,
            sharepoint_folder_url=s.sharepoint_folder_url,
            local_folder=s.local_folder,
            output_dir=s.output_dir,
            existing_article_mode=s.existing_article_mode,
            dry_run=dry_run,
        )


@dataclass
class ProgressEvent:
    """A structured event emitted during a load."""
    kind: str  # info | progress | warning | error | file_done | summary
    message: str = ""
    # For progress/file_done events:
    current: int = 0
    total: int = 0
    file_name: str = ""
    status: dict = field(default_factory=dict)  # e.g. {"content": "yes", "html": "saved", "kb": "created"}


@dataclass
class LoadResult:
    """The outcome of a load operation."""
    converted: int = 0
    created: int = 0
    updated: int = 0
    skipped: int = 0
    errors: int = 0
    log_path: Optional[Path] = None
    run_log_path: Optional[Path] = None

    @property
    def success(self) -> bool:
        return self.errors == 0


# Type alias for progress callback
ProgressCallback = Callable[[ProgressEvent], None]


def _enumerate_local_docx(folder: str):
    """Recursively find all Word files (.docx and .doc) in a local folder."""
    root = Path(folder)
    if not root.is_dir():
        raise ValueError(f"Local folder not found: {folder}")

    files = []
    for path in sorted(root.rglob("*")):
        if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
            continue
        if path.name.startswith("~$"):
            continue
        rel = path.relative_to(root)
        rel_dir = str(rel.parent) if rel.parent != Path(".") else ""
        files.append(DocxFile(
            name=path.name,
            relative_path=rel_dir,
            source_display=str(rel),
            local_path=path,
        ))
    return files


def setup_file_logging(output_dir: str, verbose: bool = False) -> Path:
    """Configure file-based logging for a load run. Returns the log file path."""
    from datetime import datetime
    log_dir = Path(output_dir)
    log_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = log_dir / f"kb_loader_{timestamp}.log"

    root_logger = logging.getLogger()
    # Remove any existing file handlers from previous runs
    for handler in list(root_logger.handlers):
        if isinstance(handler, logging.FileHandler):
            root_logger.removeHandler(handler)
            handler.close()

    root_logger.setLevel(logging.DEBUG)
    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setLevel(logging.DEBUG if verbose else logging.INFO)
    file_handler.setFormatter(logging.Formatter(
        "%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    ))
    root_logger.addHandler(file_handler)
    return log_file


def run_load(
    config: LoadConfig,
    auth: AuthClient,
    on_progress: Optional[ProgressCallback] = None,
) -> LoadResult:
    """Execute a knowledge base load operation.

    Args:
        config: Resolved configuration (input source, output, etc.).
        auth: Authenticated client. Caller is responsible for ensuring auth works.
        on_progress: Optional callback that receives ProgressEvent objects.

    Returns:
        LoadResult with counters and log paths.
    """
    def emit(event: ProgressEvent):
        if on_progress:
            try:
                on_progress(event)
            except Exception as e:
                logger.warning(f"Progress callback error: {e}")

    result = LoadResult()
    log_file = setup_file_logging(config.output_dir)
    result.log_path = log_file

    emit(ProgressEvent("info", f"Detail log: {log_file}"))

    # Step 1: Enumerate files
    emit(ProgressEvent("info", "Scanning for Word files..."))
    try:
        if config.input_mode == "local":
            files = _enumerate_local_docx(config.local_folder)
            sp_client = None
        else:
            from kb_loader.sharepoint_client import SharePointClient
            sp_client = SharePointClient(auth)
            sp_files = sp_client.enumerate_docx_files(config.sharepoint_folder_url)
            files = [
                DocxFile(
                    name=f.name,
                    relative_path=f.relative_path,
                    source_display=f"{f.relative_path}/{f.name}" if f.relative_path else f.name,
                    sp_file=f,
                )
                for f in sp_files
            ]
    except Exception as e:
        from kb_loader.sharepoint_client import SharingLinkResolutionError
        if isinstance(e, SharingLinkResolutionError):
            # Let the caller (GUI) handle this with a recovery dialog
            raise
        emit(ProgressEvent("error", f"Failed to enumerate files: {e}"))
        result.errors = 1
        return result

    if not files:
        emit(ProgressEvent("info", "No Word files found. Nothing to do."))
        return result

    total = len(files)
    emit(ProgressEvent("info", f"Found {total} Word file(s) to process."))

    # Step 2: Initialize Dataverse client (live mode only)
    dv_client = None
    if not config.dry_run:
        from kb_loader.config import Config as LegacyConfig
        legacy_cfg = LegacyConfig(
            dataverse_url=config.dataverse_url,
            output_dir=config.output_dir,
            existing_article_mode=config.existing_article_mode,
            sharepoint_folder_url=config.sharepoint_folder_url,
            local_folder=config.local_folder,
        )
        dv_client = DataverseClient(auth, legacy_cfg)

    # Step 3: Snapshot KB article counts
    run_log = RunLog()
    if dv_client:
        try:
            emit(ProgressEvent("info", "Fetching KB article counts..."))
            pre_counts = dv_client.get_article_counts_by_status()
            run_log.set_pre_counts(pre_counts)
        except Exception as e:
            emit(ProgressEvent("warning", f"Could not fetch pre-run counts: {e}"))

    # Step 4: Process each file
    for i, doc_file in enumerate(files, 1):
        log_entry = {
            "file_name": doc_file.name,
            "folder_path": doc_file.relative_path,
        }
        status = {}

        emit(ProgressEvent(
            "progress", f"Processing: {doc_file.source_display}",
            current=i, total=total, file_name=doc_file.name,
        ))
        logger.info(f"\n[{i}/{total}] Processing: {doc_file.source_display}")

        try:
            # Read file content
            if doc_file.local_path:
                docx_bytes = doc_file.local_path.read_bytes()
            else:
                docx_bytes = sp_client.download_file(doc_file.sp_file)
            log_entry["file_size"] = len(docx_bytes)

            # Convert to HTML
            html, warnings = convert_to_html(docx_bytes, doc_file.name)
            for w in warnings:
                logger.warning(f"Conversion warning: {w}")
            result.converted += 1

            has_content = bool(html and html.strip())
            log_entry["has_content"] = has_content
            status["content"] = "yes" if has_content else "EMPTY"

            # Save HTML locally only if there is content
            if has_content:
                html_path = save_html_file(
                    html, config.output_dir, doc_file.relative_path, doc_file.name
                )
                log_entry["html_saved"] = True
                status["html"] = "saved"
            else:
                log_entry["html_saved"] = False
                status["html"] = "skipped"

            if config.dry_run:
                log_entry["kb_action"] = "Dry Run"
                result.skipped += 1
                status["kb"] = "dry run"
                run_log.add_entry(**log_entry)
                emit(ProgressEvent(
                    "file_done", f"{doc_file.name}",
                    current=i, total=total, file_name=doc_file.name, status=status,
                ))
                continue

            # Look up article by title
            title = Path(doc_file.name).stem
            source_path = doc_file.source_display
            existing = dv_client.find_existing_article(title)

            if existing and config.existing_article_mode == "skip":
                log_entry["kb_action"] = "Skipped"
                log_entry["article_id"] = existing["knowledgearticleid"]
                result.skipped += 1
                status["kb"] = "skipped (exists)"
            elif existing and config.existing_article_mode == "update":
                article_id = existing["knowledgearticleid"]
                dv_client.update_article_content(article_id, title, html, source_path)
                dv_client.publish_article(article_id)
                log_entry["kb_action"] = "Updated"
                log_entry["article_id"] = article_id
                log_entry["published"] = True
                result.updated += 1
                status["kb"] = "updated"
            else:
                article_id = dv_client.create_article(title, html, source_path)
                dv_client.publish_article(article_id)
                log_entry["kb_action"] = "Created"
                log_entry["article_id"] = article_id
                log_entry["published"] = True
                result.created += 1
                status["kb"] = "created"

            run_log.add_entry(**log_entry)
            emit(ProgressEvent(
                "file_done", f"{doc_file.name}",
                current=i, total=total, file_name=doc_file.name, status=status,
            ))

        except Exception as e:
            logger.exception(f"Error processing {doc_file.source_display}: {e}")
            log_entry["error"] = str(e)
            log_entry["kb_action"] = "Error"
            run_log.add_entry(**log_entry)
            result.errors += 1
            status["kb"] = f"ERROR: {e}"
            emit(ProgressEvent(
                "file_done", f"{doc_file.name}: {e}",
                current=i, total=total, file_name=doc_file.name, status=status,
            ))

    # Step 5: Snapshot post-run KB counts
    if dv_client:
        try:
            post_counts = dv_client.get_article_counts_by_status()
            run_log.set_post_counts(post_counts)
        except Exception as e:
            emit(ProgressEvent("warning", f"Could not fetch post-run counts: {e}"))

    # Step 6: Save run log
    result.run_log_path = run_log.save(config.output_dir)

    # Step 7: Summary
    emit(ProgressEvent(
        "summary",
        f"Converted: {result.converted}, Created: {result.created}, "
        f"Updated: {result.updated}, Skipped: {result.skipped}, Errors: {result.errors}",
    ))

    return result
