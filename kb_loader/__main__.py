"""CLI entry point for D365 Knowledge Base Loader.

Usage:
    python -m kb_loader                                # launch GUI
    python -m kb_loader --gui                          # launch GUI (explicit)
    python -m kb_loader --local-folder "C:\\docs"      # CLI mode
    python -m kb_loader --sharepoint-url "https://..."  # CLI mode
    python -m kb_loader --kb-status                    # status check
    python -m kb_loader --help
"""

import argparse
import logging
import sys
from pathlib import Path

from kb_loader.auth import AuthClient
from kb_loader.dataverse_client import DataverseClient
from kb_loader.service import LoadConfig, ProgressEvent, run_load
from kb_loader.settings import Settings, load_settings


def setup_console_logging():
    """Console handler that only shows warnings and above (progress prints itself)."""
    root_logger = logging.getLogger()
    if not root_logger.handlers:
        root_logger.setLevel(logging.DEBUG)
    # Remove pre-existing console handlers
    for handler in list(root_logger.handlers):
        if isinstance(handler, logging.StreamHandler) and not isinstance(handler, logging.FileHandler):
            root_logger.removeHandler(handler)

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.WARNING)
    console_handler.setFormatter(logging.Formatter("%(message)s"))
    root_logger.addHandler(console_handler)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="kb_loader",
        description="Load Word documents into D365 Knowledge Base articles.",
    )
    parser.add_argument(
        "--gui",
        action="store_true",
        help="Launch the graphical interface (default if no other args).",
    )
    source = parser.add_mutually_exclusive_group()
    source.add_argument(
        "--local-folder",
        help="Local folder containing Word files (.docx and .doc).",
    )
    source.add_argument(
        "--sharepoint-url",
        help="SharePoint folder URL containing Word files.",
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


def _print_progress(event: ProgressEvent):
    """Render a ProgressEvent to the console for CLI mode."""
    if event.kind == "info":
        print(event.message)
    elif event.kind == "warning":
        print(f"WARNING: {event.message}")
    elif event.kind == "error":
        print(f"ERROR: {event.message}")
    elif event.kind == "file_done":
        status_str = ", ".join(f"{k}: {v}" for k, v in event.status.items())
        print(f"  [{event.current}/{event.total}] {event.file_name}  →  {status_str}")
    elif event.kind == "summary":
        print("\n" + "=" * 60)
        print(event.message)
        print("=" * 60)


def cmd_kb_status(settings: Settings):
    """Show current KB article counts and exit."""
    if not settings.dataverse_url:
        print("ERROR: Dataverse URL is not configured.", file=sys.stderr)
        sys.exit(1)

    auth = AuthClient(settings.azure_client_id, settings.azure_tenant_id)

    from kb_loader.config import Config as LegacyConfig
    legacy_cfg = LegacyConfig(
        dataverse_url=settings.dataverse_url,
        output_dir=settings.output_dir,
        existing_article_mode=settings.existing_article_mode,
    )
    dv_client = DataverseClient(auth, legacy_cfg)
    try:
        counts = dv_client.get_article_counts_by_status()
    except Exception as e:
        print(f"ERROR: Failed to fetch article counts: {e}", file=sys.stderr)
        sys.exit(1)

    print("\nKB Article Summary")
    print("=" * 35)
    for status in sorted(counts.keys() - {"Total"}):
        print(f"  {status:<20} {counts[status]:>6}")
    print("-" * 35)
    print(f"  {'Total':<20} {counts.get('Total', 0):>6}")
    print("=" * 35)


def main():
    args = parse_args()

    # Default to GUI when no CLI flags are given (other than --verbose)
    no_cli_flags = (
        not args.local_folder
        and not args.sharepoint_url
        and not args.dry_run
        and not args.kb_status
        and not args.output_dir
        and not args.existing
    )
    if args.gui or no_cli_flags:
        try:
            from kb_loader.gui import launch_gui
        except ImportError as e:
            print(f"GUI dependencies missing: {e}", file=sys.stderr)
            sys.exit(1)
        launch_gui()
        return

    # CLI mode
    setup_console_logging()
    settings = load_settings()

    # Apply CLI overrides
    if args.local_folder:
        settings.local_folder = args.local_folder
        settings.sharepoint_folder_url = ""
    if args.sharepoint_url:
        settings.sharepoint_folder_url = args.sharepoint_url
        settings.local_folder = ""
    if args.output_dir:
        settings.output_dir = args.output_dir
    if args.existing:
        settings.existing_article_mode = args.existing

    if args.kb_status:
        cmd_kb_status(settings)
        return

    # Validate before running
    ok, error = settings.is_valid_for_run()
    if not ok:
        print(f"ERROR: {error}", file=sys.stderr)
        sys.exit(1)

    config = LoadConfig.from_settings(settings, dry_run=args.dry_run)

    print("=" * 60)
    print("D365 Knowledge Base Loader")
    print("=" * 60)
    mode_label = "DRY RUN" if args.dry_run else "LIVE"
    print(f"  Mode: {mode_label}  |  Existing: {config.existing_article_mode}")
    if config.input_mode == "local":
        print(f"  Source: {config.local_folder}")
    else:
        print(f"  Source: {config.sharepoint_folder_url}")
    print("=" * 60)

    auth = AuthClient(settings.azure_client_id, settings.azure_tenant_id)

    # Pre-flight auth check (live mode only)
    if not args.dry_run:
        print("\nAuthenticating...")
        try:
            auth.get_dataverse_token(config.dataverse_url)
            if config.input_mode == "sharepoint":
                auth.get_graph_token()
            print("  Authentication successful.\n")
        except RuntimeError as e:
            print(f"\n  Authentication failed: {e}", file=sys.stderr)
            sys.exit(1)

    result = run_load(config, auth, on_progress=_print_progress)

    print(f"\nDetail log: {result.log_path}")
    if result.run_log_path:
        print(f"Run log:    {result.run_log_path}")

    if result.errors > 0:
        sys.exit(1)


if __name__ == "__main__":
    main()
