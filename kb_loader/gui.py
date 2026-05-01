"""Modern Tkinter GUI for D365 Knowledge Base Loader.

Built on ttkbootstrap for a clean, modern look across Windows, macOS, and Linux.

Single-window app for unskilled users:
  - Sign in / sign out (status indicator)
  - Pick input source (SharePoint URL or local folder)
  - Set Dataverse URL, output folder, existing-article mode
  - Test Connection / Dry Run / Run buttons
  - Live progress with per-file status
  - Open log / open output folder buttons after run

The GUI runs the load operation on a worker thread and pumps progress events
back to the UI via a queue + tkinter's `after()` polling.
"""

import logging
import os
import platform
import queue
import subprocess
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog
from typing import Optional

import ttkbootstrap as ttk
from ttkbootstrap.constants import (
    BOTH,
    DANGER,
    DISABLED,
    END,
    INFO,
    LEFT,
    NORMAL,
    PRIMARY,
    RIGHT,
    SECONDARY,
    SUCCESS,
    WARNING,
    X,
    Y,
    YES,
)
from ttkbootstrap.dialogs import Messagebox
from ttkbootstrap.scrolled import ScrolledText

from kb_loader.auth import AuthClient
from kb_loader.dataverse_client import DataverseClient
from kb_loader.service import LoadConfig, ProgressEvent, run_load
from kb_loader.settings import Settings, load_settings, save_settings

logger = logging.getLogger(__name__)

APP_TITLE = "D365 Knowledge Base Loader"
APP_SUBTITLE = "Bulk-load Word documents into Dynamics 365 Knowledge articles"
DEFAULT_THEME = "cosmo"  # clean, modern, blue accent
WINDOW_WIDTH = 1060
WINDOW_HEIGHT = 920

# ── Platform-aware font selection ─────────────────────────────────────
_SYSTEM = platform.system()

if _SYSTEM == "Windows":
    FONT_FAMILY = "Segoe UI"
    FONT_FAMILY_BOLD = "Segoe UI Semibold"
    FONT_FAMILY_MONO = "Cascadia Mono"
    FONT_FAMILY_SYMBOL = "Segoe UI Symbol"
elif _SYSTEM == "Darwin":  # macOS
    FONT_FAMILY = "SF Pro Text"  # falls back to system default if unavailable
    FONT_FAMILY_BOLD = "SF Pro Text"  # use weight option for bold
    FONT_FAMILY_MONO = "Menlo"
    FONT_FAMILY_SYMBOL = "Apple Symbols"
else:  # Linux/other
    FONT_FAMILY = "DejaVu Sans"
    FONT_FAMILY_BOLD = "DejaVu Sans"
    FONT_FAMILY_MONO = "DejaVu Sans Mono"
    FONT_FAMILY_SYMBOL = "DejaVu Sans"


def font_regular(size: int = 10) -> tuple:
    return (FONT_FAMILY, size)


def font_bold(size: int = 10) -> tuple:
    # On Mac, "SF Pro Text" doesn't have a "Semibold" variant in the family name —
    # use weight tuple instead. Tk's tuple form is (family, size, *modifiers).
    if _SYSTEM == "Darwin":
        return (FONT_FAMILY_BOLD, size, "bold")
    return (FONT_FAMILY_BOLD, size)


def font_mono(size: int = 10, bold: bool = False) -> tuple:
    if bold:
        return (FONT_FAMILY_MONO, size, "bold")
    return (FONT_FAMILY_MONO, size)


def font_symbol(size: int = 14) -> tuple:
    return (FONT_FAMILY_SYMBOL, size)


def _open_path_in_explorer(path: Path):
    """Open a file or folder in the OS file manager."""
    path = Path(path).resolve()
    try:
        if _SYSTEM == "Windows":
            os.startfile(str(path))  # type: ignore[attr-defined]
        elif _SYSTEM == "Darwin":
            subprocess.run(["open", str(path)], check=False)
        else:
            subprocess.run(["xdg-open", str(path)], check=False)
    except Exception as e:
        Messagebox.show_error(f"Could not open:\n{path}\n\n{e}", "Error")


class KBLoaderGUI:
    """Main GUI window."""

    def __init__(self, root: ttk.Window):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")
        self.root.minsize(960, 820)

        self.settings = load_settings()
        self.auth: Optional[AuthClient] = None
        self.event_queue: queue.Queue = queue.Queue()
        self.worker_thread: Optional[threading.Thread] = None
        self.last_log_path: Optional[Path] = None
        self.last_run_log_path: Optional[Path] = None
        self._device_code_dialog: Optional[tk.Toplevel] = None

        self._build_ui()
        self._refresh_auth_status()

        # Start the event pump
        self.root.after(100, self._drain_event_queue)

    # ── UI construction ────────────────────────────────────────────────

    def _build_ui(self):
        # Main container
        container = ttk.Frame(self.root, padding=(20, 16, 20, 12))
        container.pack(fill=BOTH, expand=YES)

        # ── Hero header ───────────────────────────────────────────────
        header = ttk.Frame(container)
        header.pack(fill=X, pady=(0, 14))

        title_lbl = ttk.Label(
            header, text=APP_TITLE,
            font=font_bold(18),
        )
        title_lbl.pack(anchor="w")

        subtitle_lbl = ttk.Label(
            header, text=APP_SUBTITLE,
            font=font_regular(10),
            bootstyle="secondary",
        )
        subtitle_lbl.pack(anchor="w", pady=(2, 0))

        ttk.Separator(container, orient="horizontal").pack(fill=X, pady=(0, 14))

        # Layout: Account (row 0) | Settings (row 1) | Actions (row 2) | Progress (row 3)
        # Only the Progress card should expand vertically.
        body = ttk.Frame(container)
        body.pack(fill=BOTH, expand=YES)
        body.columnconfigure(0, weight=1)
        body.rowconfigure(0, weight=0)  # Account — fixed
        body.rowconfigure(1, weight=0)  # Settings — fixed
        body.rowconfigure(2, weight=0)  # Actions — fixed
        body.rowconfigure(3, weight=1)  # Progress — expands

        # ── Account card ──────────────────────────────────────────────
        account_card = ttk.Labelframe(body, text="  Account  ", padding=14, bootstyle=PRIMARY)
        account_card.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        account_card.columnconfigure(0, weight=1)

        account_left = ttk.Frame(account_card)
        account_left.grid(row=0, column=0, sticky="w")

        # Status icon + label
        self.auth_icon_var = tk.StringVar(value="○")
        self.auth_icon = ttk.Label(
            account_left, textvariable=self.auth_icon_var,
            font=font_symbol(14),
            bootstyle="secondary",
        )
        self.auth_icon.pack(side=LEFT, padx=(0, 8))

        auth_text_frame = ttk.Frame(account_left)
        auth_text_frame.pack(side=LEFT)

        self.auth_status_var = tk.StringVar(value="Not signed in")
        ttk.Label(
            auth_text_frame, textvariable=self.auth_status_var,
            font=font_bold(11),
        ).pack(anchor="w")

        self.auth_method_var = tk.StringVar(value="")
        ttk.Label(
            auth_text_frame, textvariable=self.auth_method_var,
            font=font_regular(9),
            bootstyle="secondary",
        ).pack(anchor="w")

        account_right = ttk.Frame(account_card)
        account_right.grid(row=0, column=1, sticky="e")
        self.signin_btn = ttk.Button(
            account_right, text="Sign In", bootstyle=PRIMARY,
            command=self._on_signin, width=10,
        )
        self.signin_btn.pack(side=LEFT, padx=(0, 6))
        self.signout_btn = ttk.Button(
            account_right, text="Sign Out", bootstyle=(SECONDARY, "outline"),
            command=self._on_signout, width=10,
        )
        self.signout_btn.pack(side=LEFT)

        # ── Settings card ────────────────────────────────────────────
        settings_card = ttk.Labelframe(body, text="  Settings  ", padding=14)
        settings_card.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        settings_card.columnconfigure(1, weight=1)

        # Dataverse URL
        ttk.Label(
            settings_card, text="Dataverse URL",
            font=font_bold(10),
        ).grid(row=0, column=0, sticky="w", padx=(0, 12), pady=(0, 2))
        self.dataverse_var = tk.StringVar(value=self.settings.dataverse_url)
        ttk.Entry(
            settings_card, textvariable=self.dataverse_var,
            font=font_regular(10),
        ).grid(row=0, column=1, columnspan=2, sticky="ew", pady=(0, 2))
        ttk.Label(
            settings_card,
            text="e.g. https://your-org.crm.dynamics.com",
            font=font_regular(9),
            bootstyle="secondary",
        ).grid(row=1, column=1, columnspan=2, sticky="w", pady=(0, 10))

        # Source picker
        ttk.Label(
            settings_card, text="Source",
            font=font_bold(10),
        ).grid(row=2, column=0, sticky="nw", padx=(0, 12), pady=(0, 2))

        source_frame = ttk.Frame(settings_card)
        source_frame.grid(row=2, column=1, columnspan=2, sticky="ew", pady=(0, 10))
        source_frame.columnconfigure(1, weight=1)

        self.source_mode_var = tk.StringVar(
            value="local" if self.settings.local_folder else "sharepoint"
        )

        # SharePoint row
        ttk.Radiobutton(
            source_frame, text="SharePoint folder URL",
            variable=self.source_mode_var, value="sharepoint",
            command=self._on_source_mode_change,
            bootstyle="primary",
        ).grid(row=0, column=0, sticky="w", padx=(0, 10), pady=(0, 2))

        self.sharepoint_var = tk.StringVar(value=self.settings.sharepoint_folder_url)
        self.sharepoint_entry = ttk.Entry(
            source_frame, textvariable=self.sharepoint_var,
            font=font_regular(10),
        )
        self.sharepoint_entry.grid(row=0, column=1, sticky="ew", pady=(0, 2))

        self.sp_help_btn = ttk.Button(
            source_frame, text="?",
            command=self._show_sharepoint_help,
            bootstyle=(INFO, "outline"),
            width=3,
        )
        self.sp_help_btn.grid(row=0, column=2, padx=(8, 0), pady=(0, 2))

        # Local folder row
        ttk.Radiobutton(
            source_frame, text="Local folder",
            variable=self.source_mode_var, value="local",
            command=self._on_source_mode_change,
            bootstyle="primary",
        ).grid(row=1, column=0, sticky="w", padx=(0, 10), pady=(4, 0))

        self.local_folder_var = tk.StringVar(value=self.settings.local_folder)
        self.local_folder_entry = ttk.Entry(
            source_frame, textvariable=self.local_folder_var,
            font=font_regular(10),
        )
        self.local_folder_entry.grid(row=1, column=1, sticky="ew", padx=(0, 8), pady=(4, 0))
        self.local_browse_btn = ttk.Button(
            source_frame, text="Browse…",
            command=self._browse_local_folder,
            bootstyle=(SECONDARY, "outline"),
            width=10,
        )
        self.local_browse_btn.grid(row=1, column=2, pady=(4, 0))

        # Output folder + Browse + helper inline
        ttk.Label(
            settings_card, text="Output folder",
            font=font_bold(10),
        ).grid(row=3, column=0, sticky="w", padx=(0, 12), pady=(0, 2))
        self.output_var = tk.StringVar(value=self.settings.output_dir)
        ttk.Entry(
            settings_card, textvariable=self.output_var,
            font=font_regular(10),
        ).grid(row=3, column=1, sticky="ew", padx=(0, 8), pady=(0, 2))
        ttk.Button(
            settings_card, text="Browse…",
            command=self._browse_output_folder,
            bootstyle=(SECONDARY, "outline"),
            width=10,
        ).grid(row=3, column=2, sticky="e", pady=(0, 2))
        ttk.Label(
            settings_card,
            text="HTML files (in subfolders), detail .log, and Excel run log all land here",
            font=font_regular(9),
            bootstyle="secondary",
        ).grid(row=4, column=1, columnspan=2, sticky="w", pady=(0, 10))

        # Existing article mode + helper text inline (same row)
        ttk.Label(
            settings_card, text="If article exists",
            font=font_bold(10),
        ).grid(row=5, column=0, sticky="w", padx=(0, 12), pady=(0, 0))
        existing_row = ttk.Frame(settings_card)
        existing_row.grid(row=5, column=1, columnspan=2, sticky="ew", pady=(0, 0))
        self.existing_var = tk.StringVar(value=self.settings.existing_article_mode)
        existing_combo = ttk.Combobox(
            existing_row, textvariable=self.existing_var,
            values=["skip", "update", "duplicate"],
            state="readonly", width=12,
            font=font_regular(10),
        )
        existing_combo.pack(side=LEFT)
        ttk.Label(
            existing_row,
            text="  skip = leave existing  ·  update = overwrite  ·  duplicate = create another",
            font=font_regular(9),
            bootstyle="secondary",
        ).pack(side=LEFT, padx=(8, 0))

        # Save settings link
        save_frame = ttk.Frame(settings_card)
        save_frame.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(8, 0))
        ttk.Button(
            save_frame, text="Save settings",
            command=self._save_settings,
            bootstyle=(INFO, "link"),
        ).pack(side=RIGHT)

        self._on_source_mode_change()

        # ── Action bar ────────────────────────────────────────────────
        action_card = ttk.Frame(body)
        action_card.grid(row=2, column=0, sticky="ew", pady=(0, 12))

        action_left = ttk.Frame(action_card)
        action_left.pack(side=LEFT)

        self.test_btn = ttk.Button(
            action_left, text="Test Connection",
            command=self._on_test_connection,
            bootstyle=(INFO, "outline"),
            width=18,
        )
        self.test_btn.pack(side=LEFT, padx=(0, 8))

        self.kb_status_btn = ttk.Button(
            action_left, text="KB Status",
            command=self._on_kb_status,
            bootstyle=(SECONDARY, "outline"),
            width=12,
        )
        self.kb_status_btn.pack(side=LEFT, padx=(0, 8))

        action_right = ttk.Frame(action_card)
        action_right.pack(side=RIGHT)

        self.dryrun_btn = ttk.Button(
            action_right, text="Dry Run",
            command=lambda: self._start_run(dry_run=True),
            bootstyle=WARNING,
            width=14,
        )
        self.dryrun_btn.pack(side=LEFT, padx=(0, 8))

        self.run_btn = ttk.Button(
            action_right, text="▶  Run",
            command=lambda: self._start_run(dry_run=False),
            bootstyle=SUCCESS,
            width=14,
        )
        self.run_btn.pack(side=LEFT)

        # ── Progress card ─────────────────────────────────────────────
        progress_card = ttk.Labelframe(body, text="  Progress  ", padding=14)
        progress_card.grid(row=3, column=0, sticky="nsew", pady=(0, 0))
        progress_card.columnconfigure(0, weight=1)
        progress_card.rowconfigure(2, weight=1)

        # Progress bar with label
        bar_frame = ttk.Frame(progress_card)
        bar_frame.grid(row=0, column=0, sticky="ew", pady=(0, 4))
        bar_frame.columnconfigure(0, weight=1)

        self.progress_bar = ttk.Progressbar(
            bar_frame, mode="determinate", maximum=100, value=0,
            bootstyle=(SUCCESS, "striped"),
        )
        self.progress_bar.grid(row=0, column=0, sticky="ew", padx=(0, 12))

        self.progress_label_var = tk.StringVar(value="Ready")
        ttk.Label(
            bar_frame, textvariable=self.progress_label_var,
            font=font_bold(10),
            width=14, anchor="e",
        ).grid(row=0, column=1, sticky="e")

        # Output utility row
        util_frame = ttk.Frame(progress_card)
        util_frame.grid(row=1, column=0, sticky="ew", pady=(8, 8))
        ttk.Label(
            util_frame, text="Live output",
            font=font_bold(10),
        ).pack(side=LEFT)
        self.open_log_btn = ttk.Button(
            util_frame, text="Open detail log",
            command=self._open_last_log,
            bootstyle=(SECONDARY, "link"),
            state=DISABLED,
        )
        self.open_log_btn.pack(side=RIGHT, padx=(8, 0))
        ttk.Button(
            util_frame, text="Open output folder",
            command=self._open_output_folder,
            bootstyle=(SECONDARY, "link"),
        ).pack(side=RIGHT)
        self.clear_log_btn = ttk.Button(
            util_frame, text="Clear",
            command=self._clear_log,
            bootstyle=(SECONDARY, "link"),
        )
        self.clear_log_btn.pack(side=RIGHT)

        # Log text area
        self.log_text = ScrolledText(
            progress_card,
            wrap="word",
            height=12,
            autohide=True,
            font=font_mono(10),
            padding=8,
        )
        self.log_text.grid(row=2, column=0, sticky="nsew")

        # Color tags for log entries
        self.log_text.text.tag_configure("info", foreground="#1f6feb")
        self.log_text.text.tag_configure("warning", foreground="#9a6700")
        self.log_text.text.tag_configure("error", foreground="#cf222e")
        self.log_text.text.tag_configure("success", foreground="#1a7f37")
        self.log_text.text.tag_configure("muted", foreground="#656d76")
        self.log_text.text.tag_configure("bold", font=font_mono(10, bold=True))

        # ── Status bar ────────────────────────────────────────────────
        status_bar = ttk.Frame(container)
        status_bar.pack(fill=X, pady=(8, 0))
        ttk.Separator(status_bar, orient="horizontal").pack(fill=X, pady=(0, 4))

        status_inner = ttk.Frame(status_bar)
        status_inner.pack(fill=X)

        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(
            status_inner, textvariable=self.status_var,
            font=font_regular(9),
            bootstyle="secondary",
        ).pack(side=LEFT)

        ttk.Label(
            status_inner, text=f"Theme: {DEFAULT_THEME}",
            font=font_regular(9),
            bootstyle="secondary",
        ).pack(side=RIGHT)

    # ── UI helpers ─────────────────────────────────────────────────────

    def _on_source_mode_change(self):
        """Enable/disable the SharePoint vs Local fields based on selection."""
        mode = self.source_mode_var.get()
        if mode == "sharepoint":
            self.sharepoint_entry.configure(state=NORMAL)
            self.sp_help_btn.configure(state=NORMAL)
            self.local_folder_entry.configure(state=DISABLED)
            self.local_browse_btn.configure(state=DISABLED)
        else:
            self.sharepoint_entry.configure(state=DISABLED)
            self.sp_help_btn.configure(state=DISABLED)
            self.local_folder_entry.configure(state=NORMAL)
            self.local_browse_btn.configure(state=NORMAL)

    def _browse_local_folder(self):
        folder = filedialog.askdirectory(title="Pick a folder with Word files")
        if folder:
            self.local_folder_var.set(folder)

    def _browse_output_folder(self):
        folder = filedialog.askdirectory(title="Pick an output folder")
        if folder:
            self.output_var.set(folder)

    def _show_sharepoint_help(self):
        """Show a wizard-style dialog explaining how to find a SharePoint folder URL."""
        dlg = tk.Toplevel(self.root)
        dlg.title("How do I get my SharePoint folder URL?")
        dlg.transient(self.root)
        dlg.resizable(False, False)

        # Center on parent
        self.root.update_idletasks()
        px = self.root.winfo_rootx()
        py = self.root.winfo_rooty()
        pw = self.root.winfo_width()
        ph = self.root.winfo_height()
        dw, dh = 620, 480
        dlg.geometry(f"{dw}x{dh}+{px + (pw - dw) // 2}+{py + (ph - dh) // 2}")

        body = ttk.Frame(dlg, padding=24)
        body.pack(fill=BOTH, expand=YES)

        ttk.Label(
            body, text="How do I get my SharePoint folder URL?",
            font=font_bold(14),
        ).pack(anchor="w", pady=(0, 4))

        ttk.Label(
            body,
            text="The app supports two URL types — both work fine:",
            font=font_regular(10),
            bootstyle="secondary",
        ).pack(anchor="w", pady=(0, 14))

        # Option 1: Address bar URL
        opt1 = ttk.Labelframe(body, text="  Option 1 — From the address bar (recommended)  ", padding=12)
        opt1.pack(fill=X, pady=(0, 10))
        ttk.Label(
            opt1,
            text=(
                "1. Open SharePoint in your browser\n"
                "2. Click into the folder you want to load (e.g. 'KB Articles')\n"
                "3. Copy the entire URL from the address bar\n"
                "4. Paste it back here\n\n"
                "Looks like:  https://your-tenant.sharepoint.com/sites/Site/Shared%20Documents/Folder"
            ),
            font=font_regular(10),
            justify="left",
        ).pack(anchor="w")

        # Option 2: Sharing link
        opt2 = ttk.Labelframe(body, text="  Option 2 — A sharing link (also OK!)  ", padding=12)
        opt2.pack(fill=X, pady=(0, 10))
        ttk.Label(
            opt2,
            text=(
                "1. In SharePoint, right-click the folder → Share → Copy link\n"
                "2. Paste it back here — the app resolves it automatically\n\n"
                "Looks like:  https://your-tenant.sharepoint.com/:f:/s/SiteName/Abc123..."
            ),
            font=font_regular(10),
            justify="left",
        ).pack(anchor="w")

        # Tip
        ttk.Label(
            body,
            text=(
                "💡  Tip: After pasting, click 'Test Connection' to verify the app "
                "can find Word files in that folder before doing a full run."
            ),
            font=font_regular(9),
            bootstyle="secondary",
            wraplength=560,
            justify="left",
        ).pack(anchor="w", pady=(8, 0))

        # Close button
        btn_row = ttk.Frame(body)
        btn_row.pack(fill=X, pady=(14, 0))
        ttk.Button(
            btn_row, text="Got it",
            command=dlg.destroy,
            bootstyle=PRIMARY,
            width=12,
        ).pack(side=RIGHT)

        dlg.lift()
        dlg.focus_force()

    def _show_sharing_link_recovery_dialog(self, sharing_url: str, message: str):
        """Shown when a SharePoint sharing link can't be auto-resolved.

        Opens the link in the browser and walks the user through pasting the
        canonical URL from the address bar back into the SharePoint field.
        """
        dlg = tk.Toplevel(self.root)
        dlg.title("Help me load this folder")
        dlg.transient(self.root)
        dlg.resizable(False, False)

        # Center on parent
        self.root.update_idletasks()
        px = self.root.winfo_rootx()
        py = self.root.winfo_rooty()
        pw = self.root.winfo_width()
        ph = self.root.winfo_height()
        dw, dh = 640, 540
        dlg.geometry(f"{dw}x{dh}+{px + (pw - dw) // 2}+{py + (ph - dh) // 2}")

        body = ttk.Frame(dlg, padding=24)
        body.pack(fill=BOTH, expand=YES)

        ttk.Label(
            body, text="Let's get the right URL together",
            font=font_bold(15),
        ).pack(anchor="w", pady=(0, 4))

        ttk.Label(
            body,
            text=(
                "We tried to auto-resolve your sharing link but couldn't read it "
                "with your current sign-in. No problem — there's a quick manual fix."
            ),
            font=font_regular(10),
            bootstyle="secondary",
            wraplength=580,
            justify="left",
        ).pack(anchor="w", pady=(0, 14))

        # Step 1
        step1 = ttk.Labelframe(body, text="  Step 1 — Open the link in your browser  ", padding=12)
        step1.pack(fill=X, pady=(0, 8))
        ttk.Label(
            step1,
            text=(
                "Click the button below. Your browser will open to the SharePoint folder."
            ),
            font=font_regular(10),
            justify="left",
            wraplength=560,
        ).pack(anchor="w", pady=(0, 8))

        def open_link():
            import webbrowser
            try:
                webbrowser.open(sharing_url)
                btn.configure(text="Opened — see your browser", bootstyle=SUCCESS)
            except Exception as e:
                Messagebox.show_error(f"Could not open browser: {e}", "Error")

        btn = ttk.Button(
            step1, text="Open link in browser",
            command=open_link,
            bootstyle=PRIMARY,
            width=24,
        )
        btn.pack(anchor="w")

        # Step 2
        step2 = ttk.Labelframe(body, text="  Step 2 — Copy the URL from the address bar  ", padding=12)
        step2.pack(fill=X, pady=(0, 8))
        ttk.Label(
            step2,
            text=(
                "After the SharePoint page loads, click the address bar at the top of "
                "your browser and copy the entire URL (Ctrl+L then Ctrl+C, or ⌘L then ⌘C).\n\n"
                "It will look something like:\n"
                "  https://your-tenant.sharepoint.com/sites/SiteName/Shared%20Documents/Folder…"
            ),
            font=font_regular(10),
            justify="left",
            wraplength=560,
        ).pack(anchor="w")

        # Step 3
        step3 = ttk.Labelframe(body, text="  Step 3 — Paste it back here  ", padding=12)
        step3.pack(fill=X, pady=(0, 8))
        ttk.Label(
            step3,
            text=(
                "Paste the URL into the SharePoint folder URL field (replacing the "
                "current value), then click Test Connection again."
            ),
            font=font_regular(10),
            justify="left",
            wraplength=560,
        ).pack(anchor="w")

        # Action buttons
        btn_row = ttk.Frame(body)
        btn_row.pack(fill=X, pady=(14, 0))
        ttk.Button(
            btn_row, text="Got it",
            command=dlg.destroy,
            bootstyle=PRIMARY,
            width=12,
        ).pack(side=RIGHT)
        ttk.Button(
            btn_row, text="Cancel",
            command=dlg.destroy,
            bootstyle=(SECONDARY, "outline"),
            width=12,
        ).pack(side=RIGHT, padx=(0, 8))

        # Auto-open the link to remove a step
        try:
            import webbrowser
            webbrowser.open(sharing_url)
            btn.configure(text="Opened — see your browser", bootstyle=SUCCESS)
        except Exception:
            pass

        dlg.lift()
        dlg.attributes("-topmost", True)
        dlg.after(500, lambda: dlg.attributes("-topmost", False))
        dlg.focus_force()

    def _collect_settings(self) -> Settings:
        """Build a Settings object from the current form values."""
        s = Settings(
            dataverse_url=self.dataverse_var.get().strip(),
            output_dir=self.output_var.get().strip() or "./output",
            existing_article_mode=self.existing_var.get().strip() or "skip",
            azure_client_id=self.settings.azure_client_id,
            azure_tenant_id=self.settings.azure_tenant_id,
        )
        if self.source_mode_var.get() == "sharepoint":
            s.sharepoint_folder_url = self.sharepoint_var.get().strip()
        else:
            s.local_folder = self.local_folder_var.get().strip()
        return s

    def _save_settings(self):
        s = self._collect_settings()
        try:
            path = save_settings(s)
            self.settings = s
            self._set_status(f"Settings saved to {path}")
            self._log(f"✓ Settings saved.\n", "success")
        except Exception as e:
            Messagebox.show_error(str(e), "Could not save settings")

    def _set_status(self, msg: str):
        self.status_var.set(msg)

    def _log(self, message: str, tag: str = ""):
        """Append a message to the log text area."""
        if tag:
            self.log_text.text.insert(END, message, tag)
        else:
            self.log_text.text.insert(END, message)
        self.log_text.text.see(END)

    def _clear_log(self):
        self.log_text.text.delete("1.0", END)

    def _set_buttons_enabled(self, enabled: bool):
        """Enable or disable action buttons during a run."""
        state = NORMAL if enabled else DISABLED
        for btn in (self.test_btn, self.dryrun_btn, self.run_btn,
                    self.kb_status_btn, self.signin_btn, self.signout_btn):
            btn.configure(state=state)

    def _set_auth_indicator(self, mode: str, user: Optional[str] = None, error: Optional[str] = None):
        """Update the auth icon + label.

        mode: 'signed_in' | 'signed_out' | 'error'
        """
        if mode == "signed_in" and user:
            self.auth_icon_var.set("●")
            self.auth_icon.configure(bootstyle=SUCCESS)
            self.auth_status_var.set(f"Signed in: {user}")
            method = self.auth.method if self.auth else ""
            self.auth_method_var.set(f"via {method}" if method else "")
        elif mode == "signed_out":
            self.auth_icon_var.set("○")
            self.auth_icon.configure(bootstyle=SECONDARY)
            self.auth_status_var.set("Not signed in")
            method = self.auth.method if self.auth else ""
            self.auth_method_var.set(f"will use {method} when you sign in" if method else "")
        elif mode == "error":
            self.auth_icon_var.set("✕")
            self.auth_icon.configure(bootstyle=DANGER)
            self.auth_status_var.set("Authentication unavailable")
            self.auth_method_var.set(error or "")

    # ── Authentication ────────────────────────────────────────────────

    def _ensure_auth_client(self) -> AuthClient:
        """Build an AuthClient (re-build if settings changed)."""
        if self.auth is None:
            try:
                self.auth = AuthClient(
                    self.settings.azure_client_id,
                    self.settings.azure_tenant_id,
                )
            except RuntimeError as e:
                raise RuntimeError(
                    f"{e}\n\n"
                    "Tip: Install Azure CLI and run `az login`, "
                    "or set AZURE_CLIENT_ID via the developer .env file."
                )
            # Wire up device-code flow so sign-in is visible in the GUI
            # instead of relying on a browser window that may pop up behind
            # other apps.
            self.auth.set_device_code_callback(self._on_device_code)
        return self.auth

    def _on_device_code(self, code: str, url: str):
        """Called from the auth worker thread when az emits a device code.
        Marshal to the UI thread via the event queue."""
        self.event_queue.put(("device_code", code, url))

    def _refresh_auth_status(self):
        """Update the auth indicator label."""
        try:
            auth = self._ensure_auth_client()
            user = auth.get_signed_in_user()
            if user:
                self._set_auth_indicator("signed_in", user=user)
            else:
                self._set_auth_indicator("signed_out")
        except RuntimeError as e:
            self._set_auth_indicator("error", error=str(e))

    # ── Device-code sign-in dialog ────────────────────────────────────

    def _show_device_code_dialog(self, code: str, url: str):
        """Display the device-code sign-in instructions in a prominent modal."""
        # If a previous dialog is still around, close it
        self._dismiss_device_code_dialog()

        dlg = tk.Toplevel(self.root)
        self._device_code_dialog = dlg
        dlg.title("Sign in to Microsoft")
        dlg.transient(self.root)
        dlg.resizable(False, False)
        dlg.protocol("WM_DELETE_WINDOW", self._dismiss_device_code_dialog)

        # Center on parent window
        self.root.update_idletasks()
        px = self.root.winfo_rootx()
        py = self.root.winfo_rooty()
        pw = self.root.winfo_width()
        ph = self.root.winfo_height()
        dw, dh = 520, 320
        dlg.geometry(f"{dw}x{dh}+{px + (pw - dw) // 2}+{py + (ph - dh) // 2}")

        body = ttk.Frame(dlg, padding=24)
        body.pack(fill=BOTH, expand=YES)

        ttk.Label(
            body, text="Sign in to Microsoft",
            font=font_bold(16),
        ).pack(anchor="w")

        ttk.Label(
            body,
            text="A browser window should open automatically. Sign in with your\n"
                 "Microsoft account, and enter the code below when prompted.",
            font=font_regular(10),
            bootstyle="secondary",
            justify="left",
        ).pack(anchor="w", pady=(6, 16))

        # Big code box
        code_frame = ttk.Frame(body, bootstyle="primary")
        code_frame.pack(fill=X, pady=(0, 12))
        code_label = ttk.Label(
            code_frame, text=code,
            font=font_mono(28, bold=True),
            bootstyle="inverse-primary",
            anchor="center", padding=(20, 16),
        )
        code_label.pack(fill=X)

        # Buttons row
        btn_row = ttk.Frame(body)
        btn_row.pack(fill=X, pady=(0, 8))

        def copy_code():
            self.root.clipboard_clear()
            self.root.clipboard_append(code)
            copy_btn.configure(text="Copied!")
            self.root.after(1500, lambda: copy_btn.configure(text="Copy code"))

        def open_browser():
            import webbrowser
            try:
                webbrowser.open(url)
            except Exception as e:
                logger.warning(f"Could not open browser: {e}")

        copy_btn = ttk.Button(
            btn_row, text="Copy code",
            command=copy_code,
            bootstyle=(SECONDARY, "outline"),
            width=14,
        )
        copy_btn.pack(side=LEFT, padx=(0, 8))

        ttk.Button(
            btn_row, text="Open sign-in page",
            command=open_browser,
            bootstyle=PRIMARY,
            width=22,
        ).pack(side=LEFT)

        ttk.Label(
            body,
            text=f"URL: {url}",
            font=font_regular(9),
            bootstyle="secondary",
        ).pack(anchor="w", pady=(8, 0))

        ttk.Label(
            body,
            text="This dialog will close automatically once sign-in completes.",
            font=font_regular(9),
            bootstyle="secondary",
        ).pack(anchor="w", pady=(4, 0))

        # Bring it to front
        dlg.lift()
        dlg.attributes("-topmost", True)
        dlg.after(500, lambda: dlg.attributes("-topmost", False))
        dlg.focus_force()

        # Auto-open the browser so the user doesn't have to hunt for the URL
        open_browser()

        # Also drop a line into the live log so it's not lost
        self._log(f"\nSign-in code: {code}\n", "bold")
        self._log(f"Sign-in URL:  {url}\n\n", "muted")

    def _dismiss_device_code_dialog(self):
        """Close the device-code dialog if it's open."""
        if self._device_code_dialog is not None:
            try:
                self._device_code_dialog.destroy()
            except Exception:
                pass
            self._device_code_dialog = None

    def _on_signin(self):
        """Trigger interactive sign-in by acquiring a Graph token."""
        self._set_status("Signing in — a sign-in code will appear shortly…")
        self._log("Starting sign-in…\n", "info")

        def worker():
            try:
                auth = self._ensure_auth_client()
                # Acquiring a Graph token will trigger interactive sign-in if needed
                auth.get_graph_token()
                self.event_queue.put(("auth_ok", auth.get_signed_in_user() or ""))
            except Exception as e:
                self.event_queue.put(("auth_err", str(e)))

        self._set_buttons_enabled(False)
        threading.Thread(target=worker, daemon=True).start()

    def _on_signout(self):
        confirm = Messagebox.yesno("Are you sure you want to sign out?", "Sign out")
        if confirm != "Yes":
            return
        try:
            auth = self._ensure_auth_client()
            auth.sign_out()
            self.auth = None
            self._refresh_auth_status()
            self._log("✓ Signed out.\n", "info")
            self._set_status("Signed out.")
        except Exception as e:
            Messagebox.show_error(str(e), "Sign out failed")

    # ── Test connection ───────────────────────────────────────────────

    def _on_test_connection(self):
        s = self._collect_settings()
        ok, error = s.is_valid_for_run()
        if not ok:
            Messagebox.show_warning(error, "Settings incomplete")
            return

        self._clear_log()
        self._log("Testing connection…\n", "info")
        self._set_status("Testing connection…")
        self._set_buttons_enabled(False)

        def worker():
            try:
                self.settings = s
                self.auth = None
                auth = self._ensure_auth_client()

                self.event_queue.put(("log", "  Acquiring Dataverse token…\n", "muted"))
                auth.get_dataverse_token(s.dataverse_url)

                from kb_loader.config import Config as LegacyConfig
                legacy_cfg = LegacyConfig(
                    dataverse_url=s.dataverse_url,
                    output_dir=s.output_dir,
                    existing_article_mode=s.existing_article_mode,
                )
                dv = DataverseClient(auth, legacy_cfg)
                counts = dv.get_article_counts_by_status()
                self.event_queue.put((
                    "log",
                    f"  ✓ Dataverse OK — {counts.get('Total', 0)} articles total\n",
                    "success",
                ))

                if s.input_mode == "sharepoint":
                    self.event_queue.put(("log", "  Acquiring Graph token…\n", "muted"))
                    auth.get_graph_token()
                    from kb_loader.sharepoint_client import SharePointClient
                    sp = SharePointClient(auth)
                    files = sp.enumerate_docx_files(s.sharepoint_folder_url)
                    self.event_queue.put((
                        "log",
                        f"  ✓ SharePoint OK — found {len(files)} Word file(s)\n",
                        "success",
                    ))
                elif s.input_mode == "local":
                    folder = Path(s.local_folder)
                    count = sum(
                        1 for p in folder.rglob("*")
                        if p.suffix.lower() in (".docx", ".doc") and not p.name.startswith("~$")
                    )
                    self.event_queue.put((
                        "log",
                        f"  ✓ Local folder OK — found {count} Word file(s)\n",
                        "success",
                    ))

                self.event_queue.put(("test_done", "ok"))
            except Exception as e:
                from kb_loader.sharepoint_client import SharingLinkResolutionError
                if isinstance(e, SharingLinkResolutionError):
                    self.event_queue.put(("sharing_link_error", e.sharing_url, str(e)))
                else:
                    self.event_queue.put(("test_done", f"err:{e}"))

        threading.Thread(target=worker, daemon=True).start()

    # ── KB status ─────────────────────────────────────────────────────

    def _on_kb_status(self):
        s = self._collect_settings()
        if not s.dataverse_url:
            Messagebox.show_warning("Please enter the Dataverse URL.", "Dataverse URL required")
            return
        self._set_buttons_enabled(False)
        self._log("Fetching KB article counts…\n", "info")

        def worker():
            try:
                self.settings = s
                self.auth = None
                auth = self._ensure_auth_client()

                from kb_loader.config import Config as LegacyConfig
                legacy_cfg = LegacyConfig(
                    dataverse_url=s.dataverse_url,
                    output_dir=s.output_dir,
                    existing_article_mode=s.existing_article_mode,
                )
                dv = DataverseClient(auth, legacy_cfg)
                counts = dv.get_article_counts_by_status()
                lines = ["\n  KB Article Summary\n", "  " + "─" * 35 + "\n"]
                for status in sorted(counts.keys() - {"Total"}):
                    lines.append(f"  {status:<20} {counts[status]:>6}\n")
                lines.append("  " + "─" * 35 + "\n")
                lines.append(f"  {'Total':<20} {counts.get('Total', 0):>6}\n\n")
                self.event_queue.put(("log_block", "".join(lines), "info"))
                self.event_queue.put(("status_done", "Ready"))
            except Exception as e:
                self.event_queue.put(("test_done", f"err:{e}"))

        threading.Thread(target=worker, daemon=True).start()

    # ── Run / Dry run ─────────────────────────────────────────────────

    def _start_run(self, dry_run: bool):
        s = self._collect_settings()
        ok, error = s.is_valid_for_run()
        if not ok:
            Messagebox.show_warning(error, "Settings incomplete")
            return

        if not dry_run:
            confirm = Messagebox.yesno(
                "This will publish articles to Dynamics 365.\n\n"
                f"Existing article mode: {s.existing_article_mode}\n\n"
                "Continue?",
                "Confirm run",
            )
            if confirm != "Yes":
                return

        # Save settings (so the user doesn't lose form values on close)
        try:
            save_settings(s)
            self.settings = s
        except Exception as e:
            logger.warning(f"Could not save settings before run: {e}")

        # Reset UI for run
        self._clear_log()
        self.progress_bar.configure(value=0, maximum=100)
        self.progress_label_var.set("Starting…")
        self._set_buttons_enabled(False)
        self.last_log_path = None
        self.last_run_log_path = None

        mode_label = "DRY RUN" if dry_run else "LIVE RUN"
        self._log(f"━━━ {mode_label} ━━━\n", "bold")
        self._set_status(f"{mode_label} in progress…")

        def worker():
            try:
                self.auth = None
                auth = self._ensure_auth_client()
                config = LoadConfig.from_settings(s, dry_run=dry_run)

                if not dry_run:
                    self.event_queue.put(("log", "Authenticating…\n", "info"))
                    auth.get_dataverse_token(config.dataverse_url)
                    if config.input_mode == "sharepoint":
                        auth.get_graph_token()
                    self.event_queue.put(("log", "  ✓ Authentication successful.\n\n", "success"))

                def emit(event: ProgressEvent):
                    self.event_queue.put(("progress_event", event))

                result = run_load(config, auth, on_progress=emit)
                self.event_queue.put(("run_done", result))
            except Exception as e:
                logger.exception("Run failed")
                from kb_loader.sharepoint_client import SharingLinkResolutionError
                if isinstance(e, SharingLinkResolutionError):
                    self.event_queue.put(("sharing_link_error", e.sharing_url, str(e)))
                else:
                    self.event_queue.put(("run_failed", str(e)))

        self.worker_thread = threading.Thread(target=worker, daemon=True)
        self.worker_thread.start()

    # ── Open helpers ──────────────────────────────────────────────────

    def _open_last_log(self):
        if self.last_log_path and Path(self.last_log_path).exists():
            _open_path_in_explorer(self.last_log_path)
        else:
            Messagebox.show_info("Run a load first to generate a log file.", "No log yet")

    def _open_output_folder(self):
        s = self._collect_settings()
        out = Path(s.output_dir or "./output").resolve()
        out.mkdir(parents=True, exist_ok=True)

        # Inventory what's there so the user sees what to expect
        html_count = sum(1 for p in out.rglob("*.html"))
        log_count = sum(1 for p in out.glob("kb_loader_*.log"))
        xlsx_count = sum(1 for p in out.glob("kb_loader_log_*.xlsx"))

        if html_count + log_count + xlsx_count == 0:
            self._log(
                f"Output folder: {out}\n"
                f"  (empty — files appear here after a Dry Run or Run)\n",
                "muted",
            )
        else:
            self._log(
                f"Output folder: {out}\n"
                f"  Contains: {html_count} HTML file(s), "
                f"{log_count} detail log(s), {xlsx_count} run log(s)\n",
                "muted",
            )
        self._set_status(f"Opened {out}")
        _open_path_in_explorer(out)

    # ── Event pump (drains worker queue) ──────────────────────────────

    def _drain_event_queue(self):
        """Process events posted by worker threads."""
        try:
            while True:
                event = self.event_queue.get_nowait()
                self._handle_event(event)
        except queue.Empty:
            pass
        self.root.after(100, self._drain_event_queue)

    def _handle_event(self, event: tuple):
        kind = event[0]
        if kind == "log":
            _, msg, tag = event
            self._log(msg, tag)
        elif kind == "log_block":
            _, msg, tag = event
            self._log(msg, tag)
        elif kind == "device_code":
            _, code, url = event
            self._show_device_code_dialog(code, url)
        elif kind == "auth_ok":
            user = event[1]
            self._dismiss_device_code_dialog()
            self._refresh_auth_status()
            if user:
                self._log(f"✓ Signed in as {user}\n", "success")
                self._set_status(f"Signed in as {user}")
            else:
                self._log("✓ Sign-in complete.\n", "success")
                self._set_status("Signed in.")
            self._set_buttons_enabled(True)
        elif kind == "auth_err":
            err = event[1]
            self._dismiss_device_code_dialog()
            self._log(f"✗ Sign-in failed: {err}\n", "error")
            self._set_status("Sign-in failed.")
            Messagebox.show_error(err, "Sign-in failed")
            self._set_buttons_enabled(True)
        elif kind == "sharing_link_error":
            _, sharing_url, message = event
            self._dismiss_device_code_dialog()
            self._log(f"⚠ {message}\n", "warning")
            self._set_status("Couldn't auto-resolve sharing link.")
            self._set_buttons_enabled(True)
            self._refresh_auth_status()
            self._show_sharing_link_recovery_dialog(sharing_url, message)
        elif kind == "test_done":
            outcome = event[1]
            self._dismiss_device_code_dialog()
            if outcome == "ok":
                self._log("✓ Connection test passed.\n", "success")
                self._set_status("Connection OK.")
            else:
                err = outcome[4:]
                self._log(f"✗ {err}\n", "error")
                self._set_status("Connection failed.")
                Messagebox.show_error(err, "Connection test failed")
            self._set_buttons_enabled(True)
            self._refresh_auth_status()
        elif kind == "status_done":
            self._dismiss_device_code_dialog()
            self._set_status(event[1])
            self._set_buttons_enabled(True)
            self._refresh_auth_status()
        elif kind == "progress_event":
            self._render_progress_event(event[1])
        elif kind == "run_done":
            self._dismiss_device_code_dialog()
            result = event[1]
            self._on_run_complete(result)
        elif kind == "run_failed":
            err = event[1]
            self._dismiss_device_code_dialog()
            self._log(f"\n✗ Run failed: {err}\n", "error")
            self._set_status("Run failed.")
            Messagebox.show_error(err, "Run failed")
            self._set_buttons_enabled(True)
            self._refresh_auth_status()

    def _render_progress_event(self, event: ProgressEvent):
        if event.kind == "info":
            self._log(f"{event.message}\n", "info")
        elif event.kind == "warning":
            self._log(f"⚠ {event.message}\n", "warning")
        elif event.kind == "error":
            self._log(f"✗ {event.message}\n", "error")
        elif event.kind == "progress":
            if event.total > 0:
                pct = (event.current / event.total) * 100
                self.progress_bar.configure(value=pct, maximum=100)
                self.progress_label_var.set(f"{event.current} / {event.total}")
        elif event.kind == "file_done":
            status_str = " · ".join(f"{k}: {v}" for k, v in event.status.items())
            tag = "error" if any("ERROR" in str(v) for v in event.status.values()) else ""
            self._log(f"  [{event.current:>3}/{event.total}] {event.file_name}\n", "muted")
            self._log(f"          {status_str}\n", tag)
            if event.total > 0:
                pct = (event.current / event.total) * 100
                self.progress_bar.configure(value=pct, maximum=100)
                self.progress_label_var.set(f"{event.current} / {event.total}")
        elif event.kind == "summary":
            self._log(f"\n{event.message}\n", "info")

    def _on_run_complete(self, result):
        self.progress_bar.configure(value=100)
        self.progress_label_var.set("Done")
        self.last_log_path = result.log_path
        self.last_run_log_path = result.run_log_path

        self._log("\n" + "━" * 60 + "\n", "muted")
        if result.success:
            self._log("✓ Run complete\n", "success")
        else:
            self._log(f"⚠ Run finished with {result.errors} error(s)\n", "warning")
        self._log(
            f"   Converted: {result.converted}   "
            f"Created: {result.created}   "
            f"Updated: {result.updated}   "
            f"Skipped: {result.skipped}   "
            f"Errors: {result.errors}\n",
            "info",
        )
        if result.log_path:
            self._log(f"   Detail log: {result.log_path}\n", "muted")
        if result.run_log_path:
            self._log(f"   Run log:    {result.run_log_path}\n", "muted")
        self._log("━" * 60 + "\n", "muted")

        self.open_log_btn.configure(state=NORMAL)
        self._set_buttons_enabled(True)
        if result.errors:
            self._set_status(f"Run finished with {result.errors} error(s).")
        else:
            self._set_status("Run complete.")
        self._refresh_auth_status()


def _enable_windows_dpi_awareness():
    """On Windows, opt into per-monitor DPI awareness so the GUI is crisp on HiDPI displays."""
    if _SYSTEM != "Windows":
        return
    try:
        import ctypes
        # PROCESS_PER_MONITOR_DPI_AWARE = 2
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except (AttributeError, OSError):
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def launch_gui():
    """Launch the Tkinter GUI. Blocks until the window is closed."""
    _enable_windows_dpi_awareness()

    root_logger = logging.getLogger()
    root_logger.setLevel(logging.DEBUG)

    root = ttk.Window(themename=DEFAULT_THEME)

    # On macOS, set a proper application name in the menu bar (instead of "Python")
    if _SYSTEM == "Darwin":
        try:
            root.tk.call("tk::mac::standardAboutPanel")
        except tk.TclError:
            pass

    KBLoaderGUI(root)
    root.mainloop()


if __name__ == "__main__":
    launch_gui()
