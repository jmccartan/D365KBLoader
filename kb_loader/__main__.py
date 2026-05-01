"""CLI entry point for D365 Knowledge Base Loader.

Usage:
    python -m kb_loader --local-folder "C:\\Users\\you\\OneDrive\\KB Articles"
    python -m kb_loader --sharepoint-url "https://tenant.sharepoint.com/sites/Site/Docs/Folder"
    python -m kb_loader --help
"""

import argparse
import logging
import sys
from dataclasses import dataclass
from pathlib import Path

from kb_loader.config import load_config
from kb_loader.auth import AuthClient
from kb_loader.converter import convert_to_html, save_html_file, SUPPORTED_EXTENSIONS
from kb_loader.dataverse_client import DataverseClient
from kb_loader.run_log import RunLog


@dataclass
class DocxFile:
    """A Word document (.docx or .doc) to process, from either local or SharePoint source."""
    name: str
    relative_path: str  # subfolder path relative to root
    source_display: str  # for logging

    # One of these will be set
    local_path: Path | None = None
    sp_file: object | None = None  # SharePointFile


def setup_logging(verbose: bool, output_dir: str = "./output"):
    """Configure logging: detailed output to a log file, progress summary to console."""
    # Ensure output directory exists for the log file
    log_dir = Path(output_dir)
    log_dir.mkdir(parents=True, exist_ok=True)

    from datetime import datetime
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = log_dir / f"kb_loader_{timestamp}.log"

    root_logger = logging.getLogger()
    root_logger.setLevel(logging.DEBUG)

    # File handler — all detail goes here
    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setLevel(logging.DEBUG if verbose else logging.INFO)
    file_handler.setFormatter(logging.Formatter(
        "%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    ))
    root_logger.addHandler(file_handler)

    # Console handler — only warnings+ from libraries, progress via print()
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.WARNING)
    console_handler.setFormatter(logging.Formatter("%(message)s"))
    root_logger.addHandler(console_handler)

    return log_file


def enumerate_local_docx(folder: str) -> list[DocxFile]:
    """Recursively find all Word files (.docx and .doc) in a local folder."""
    root = Path(folder)
    if not root.is_dir():
        raise ValueError(f"Local folder not found: {folder}")

    files = []
    for path in sorted(root.rglob("*")):
        if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
            continue
        # Skip temp/hidden files (e.g. ~$document.docx)
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


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="kb_loader",
        description="Load Word documents into D365 Knowledge Base articles.",
    )
    source = parser.add_mutually_exclusive_group()
    source.add_argument(
        "--local-folder",
        help="Local folder containing Word files (.docx and .doc).",
    )
    source.add_argument(
        "--sharepoint-url",
        help="SharePoint folder URL containing Word files (requires az login).",
    )
    parser.add_argument(
        "--output-dir",
        help="Local directory for saving HTML files (default: ./output).",
    )
    parser.add_argument(
        "--existing",
        choices=["skip", "update", "duplicate"],
        help="How to handle articles that already exist by title (default: skip).",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Convert files to HTML but don't upload to Dataverse.",
    )
    parser.add_argument(
        "--kb-status",
        action="store_true",
        help="Show current KB article counts by status and exit (no processing).",
    )
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="Enable verbose/debug logging.",
    )
    return parser.parse_args()


def main():
    args = parse_args()

    # Determine output dir early for log file placement
    output_dir = args.output_dir or "./output"
    log_file = setup_logging(args.verbose, output_dir)
    logger = logging.getLogger("kb_loader")

    # --kb-status: just show article counts and exit
    if args.kb_status:
        try:
            config = load_config(
                sharepoint_url="dummy",  # not needed, just satisfy validation
                output_dir=args.output_dir,
                existing_mode=args.existing,
            )
        except ValueError:
            # Fallback: load just the dataverse URL from env
            from dotenv import load_dotenv
            import os
            load_dotenv()
            dataverse_url = os.getenv("DATAVERSE_URL", "")
            if not dataverse_url:
                logger.error("DATAVERSE_URL is required. Set it in .env or environment.")
                sys.exit(1)
            from kb_loader.config import Config
            config = Config(dataverse_url=dataverse_url, output_dir="./output", existing_article_mode="skip")

        auth = AuthClient()
        dv_client = DataverseClient(auth, config)
        try:
            counts = dv_client.get_article_counts_by_status()
        except Exception as e:
            logger.error(f"Failed to fetch article counts: {e}")
            sys.exit(1)

        print("\nKB Article Summary")
        print("=" * 35)
        for status in sorted(counts.keys() - {"Total"}):
            print(f"  {status:<20} {counts[status]:>6}")
        print("-" * 35)
        print(f"  {'Total':<20} {counts.get('Total', 0):>6}")
        print("=" * 35)
        sys.exit(0)

    try:
        config = load_config(
            sharepoint_url=args.sharepoint_url,
            local_folder=args.local_folder,
            output_dir=args.output_dir,
            existing_mode=args.existing,
        )
    except ValueError as e:
        logger.error(str(e))
        sys.exit(1)

    logger.info("=" * 60)
    logger.info("D365 Knowledge Base Loader")
    logger.info("=" * 60)
    if config.input_mode == "local":
        logger.info(f"Input          : Local folder: {config.local_folder}")
    else:
        logger.info(f"Input          : SharePoint: {config.sharepoint_folder_url}")
    logger.info(f"Dataverse URL  : {config.dataverse_url}")
    logger.info(f"Output dir     : {config.output_dir}")
    logger.info(f"Existing mode  : {config.existing_article_mode}")
    logger.info(f"Dry run        : {args.dry_run}")
    logger.info("=" * 60)

    # Console banner
    print("=" * 60)
    print("D365 Knowledge Base Loader")
    print("=" * 60)
    mode_label = "DRY RUN" if args.dry_run else "LIVE"
    print(f"  Mode: {mode_label}  |  Existing: {config.existing_article_mode}")
    print(f"  Log file: {log_file}")
    print("=" * 60)

    # Initialize auth (triggers az login if needed)
    auth = AuthClient()

    # Step 1: Enumerate .docx files
    logger.info("Scanning for Word files (.docx, .doc)...")
    try:
        if config.input_mode == "local":
            files = enumerate_local_docx(config.local_folder)
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
        logger.error(f"Failed to enumerate files: {e}")
        sys.exit(1)

    if not files:
        print("No Word files found. Nothing to do.")
        logger.info("No Word files found. Nothing to do.")
        sys.exit(0)

    total_files = len(files)
    print(f"\nFound {total_files} Word file(s) to process.\n")
    logger.info(f"Found {total_files} Word file(s) to process.")

    # Initialize Dataverse client (only if not dry-run, to avoid unnecessary login)
    dv_client = None
    if not args.dry_run:
        dv_client = DataverseClient(auth, config)

        # Pre-flight: acquire tokens now so auth issues are caught before processing
        print("Authenticating...")
        try:
            auth.get_dataverse_token(config.dataverse_url)
            if config.input_mode == "sharepoint":
                auth.get_graph_token()
            print("  Authentication successful.\n")
        except RuntimeError as e:
            print(f"\n  Authentication failed: {e}")
            logger.error(f"Authentication failed: {e}")
            sys.exit(1)

    # Step 2: Snapshot KB article counts before processing
    run_log = RunLog()
    if dv_client:
        try:
            logger.info("Fetching KB article counts (before)...")
            pre_counts = dv_client.get_article_counts_by_status()
            run_log.set_pre_counts(pre_counts)
            for status, count in sorted(pre_counts.items()):
                logger.info(f"  {status}: {count}")
        except Exception as e:
            logger.warning(f"Could not fetch pre-run article counts: {e}")

    # Step 3: Process each file
    stats = {"converted": 0, "created": 0, "updated": 0, "skipped": 0, "errors": 0}

    for i, doc_file in enumerate(files, 1):
        logger.info(f"\n[{i}/{total_files}] Processing: {doc_file.source_display}")
        status_parts = []

        log_entry = {
            "file_name": doc_file.name,
            "folder_path": doc_file.relative_path,
        }

        try:
            # Read file content
            if doc_file.local_path:
                docx_bytes = doc_file.local_path.read_bytes()
                logger.info(f"  Read {len(docx_bytes):,} bytes from local file.")
            else:
                from kb_loader.sharepoint_client import SharePointClient
                logger.info(f"  Downloading from SharePoint...")
                docx_bytes = sp_client.download_file(doc_file.sp_file)

            log_entry["file_size"] = len(docx_bytes)

            # Convert to HTML
            logger.info("  Converting to HTML...")
            html, warnings = convert_to_html(docx_bytes, doc_file.name)
            if warnings:
                for w in warnings:
                    logger.warning(f"  Conversion warning: {w}")
            stats["converted"] += 1

            has_content = bool(html and html.strip())
            log_entry["has_content"] = has_content

            if has_content:
                status_parts.append("content: yes")
            else:
                status_parts.append("content: EMPTY")

            # Save HTML locally only if there is content
            if has_content:
                html_path = save_html_file(
                    html, config.output_dir, doc_file.relative_path, doc_file.name
                )
                logger.info(f"  Saved HTML: {html_path}")
                log_entry["html_saved"] = True
                status_parts.append("html: saved")
            else:
                logger.info("  No content — skipping HTML file.")
                log_entry["html_saved"] = False
                status_parts.append("html: skipped")

            if args.dry_run:
                logger.info("  [DRY RUN] Skipping Dataverse upload.")
                log_entry["kb_action"] = "Dry Run"
                stats["skipped"] += 1
                status_parts.append("KB: dry run")
                run_log.add_entry(**log_entry)
                print(f"  [{i}/{total_files}] {doc_file.name}  →  {', '.join(status_parts)}")
                continue

            # Prepare article metadata
            title = Path(doc_file.name).stem
            source_path = doc_file.source_display

            # Check for existing article
            existing = dv_client.find_existing_article(title)

            if existing and config.existing_article_mode == "skip":
                logger.info(f"  Article '{title}' already exists. Skipping.")
                log_entry["kb_action"] = "Skipped"
                log_entry["article_id"] = existing["knowledgearticleid"]
                stats["skipped"] += 1
                status_parts.append("KB: skipped (exists)")
                run_log.add_entry(**log_entry)
                print(f"  [{i}/{total_files}] {doc_file.name}  →  {', '.join(status_parts)}")
                continue

            if existing and config.existing_article_mode == "update":
                article_id = existing["knowledgearticleid"]
                logger.info(f"  Updating existing article '{title}' ({article_id})...")
                dv_client.update_article_content(article_id, title, html, source_path)
                logger.info(f"  Publishing updated article...")
                dv_client.publish_article(article_id)
                log_entry["kb_action"] = "Updated"
                log_entry["article_id"] = article_id
                log_entry["published"] = True
                stats["updated"] += 1
                status_parts.append("KB: updated")
            else:
                # Create new article
                logger.info(f"  Creating article: '{title}'...")
                article_id = dv_client.create_article(title, html, source_path)

                # Publish the article
                logger.info(f"  Publishing article...")
                dv_client.publish_article(article_id)
                log_entry["kb_action"] = "Created"
                log_entry["article_id"] = article_id
                log_entry["published"] = True
                stats["created"] += 1
                status_parts.append("KB: created")

            logger.info(f"  Done.")
            run_log.add_entry(**log_entry)

        except Exception as e:
            logger.exception(f"  Error processing {doc_file.source_display}: {e}")
            log_entry["error"] = str(e)
            log_entry["kb_action"] = "Error"
            run_log.add_entry(**log_entry)
            stats["errors"] += 1
            status_parts.append(f"ERROR: {e}")

        print(f"  [{i}/{total_files}] {doc_file.name}  →  {', '.join(status_parts)}")

    # Snapshot KB article counts after processing
    if dv_client:
        try:
            logger.info("Fetching KB article counts (after)...")
            post_counts = dv_client.get_article_counts_by_status()
            run_log.set_post_counts(post_counts)
            for status, count in sorted(post_counts.items()):
                logger.info(f"  {status}: {count}")
        except Exception as e:
            logger.warning(f"Could not fetch post-run article counts: {e}")

    # Save run log
    log_path = run_log.save(config.output_dir)

    # Summary — both console and log file
    summary_lines = [
        "",
        "=" * 60,
        "Processing complete!",
        f"  Files converted : {stats['converted']}",
        f"  Articles created: {stats['created']}",
        f"  Articles updated: {stats['updated']}",
        f"  Skipped         : {stats['skipped']}",
        f"  Errors          : {stats['errors']}",
        f"  Run log         : {log_path}",
        f"  Detail log      : {log_file}",
        "=" * 60,
    ]
    for line in summary_lines:
        logger.info(line)
        print(line)

    if stats["errors"] > 0:
        sys.exit(1)


if __name__ == "__main__":
    main()
