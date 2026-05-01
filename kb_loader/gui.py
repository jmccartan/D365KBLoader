"""Tkinter GUI for D365 Knowledge Base Loader.

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
import sys
import threading
import tkinter as tk
import webbrowser
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Optional

from kb_loader.auth import AuthClient
from kb_loader.dataverse_client import DataverseClient
from kb_loader.service import LoadConfig, ProgressEvent, run_load
from kb_loader.settings import Settings, load_settings, save_settings

logger = logging.getLogger(__name__)

APP_TITLE = "D365 Knowledge Base Loader"
WINDOW_SIZE = "880x720"


def _open_path_in_explorer(path: Path):
    """Open a file or folder in the OS file manager."""
    path = Path(path).resolve()
    system = platform.system()
    try:
        if system == "Windows":
            os.startfile(str(path))  # type: ignore[attr-defined]
        elif system == "Darwin":
            subprocess.run(["open", str(path)], check=False)
        else:
            subprocess.run(["xdg-open", str(path)], check=False)
    except Exception as e:
        messagebox.showerror("Could not open", f"{path}\n\n{e}")


class KBLoaderGUI:
    """Main GUI window."""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.minsize(720, 600)

        self.settings = load_settings()
        self.auth: Optional[AuthClient] = None
        self.event_queue: queue.Queue = queue.Queue()
        self.worker_thread: Optional[threading.Thread] = None
        self.last_log_path: Optional[Path] = None
        self.last_run_log_path: Optional[Path] = None

        self._build_ui()
        self._refresh_auth_status()

        # Start the event pump
        self.root.after(100, self._drain_event_queue)

    # ── UI construction ────────────────────────────────────────────────

    def _build_ui(self):
        # Main container with padding
        main = ttk.Frame(self.root, padding=12)
        main.pack(fill=tk.BOTH, expand=True)

        # Configure grid weights for resizing
        main.columnconfigure(0, weight=1)
        main.rowconfigure(4, weight=1)  # progress area expands

        # ── Header ────────────────────────────────────────────────────
        header = ttk.Frame(main)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        header.columnconfigure(0, weight=1)
        ttk.Label(
            header, text=APP_TITLE,
            font=("Segoe UI", 14, "bold"),
        ).grid(row=0, column=0, sticky="w")

        # ── Auth row ──────────────────────────────────────────────────
        auth_frame = ttk.LabelFrame(main, text="Account", padding=10)
        auth_frame.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        auth_frame.columnconfigure(0, weight=1)

        self.auth_status_var = tk.StringVar(value="Not signed in")
        self.auth_status_label = ttk.Label(
            auth_frame, textvariable=self.auth_status_var,
            font=("Segoe UI", 10),
        )
        self.auth_status_label.grid(row=0, column=0, sticky="w")

        auth_buttons = ttk.Frame(auth_frame)
        auth_buttons.grid(row=0, column=1, sticky="e")
        self.signin_btn = ttk.Button(auth_buttons, text="Sign In", command=self._on_signin)
        self.signin_btn.pack(side=tk.LEFT, padx=4)
        self.signout_btn = ttk.Button(auth_buttons, text="Sign Out", command=self._on_signout)
        self.signout_btn.pack(side=tk.LEFT, padx=4)

        # ── Settings ──────────────────────────────────────────────────
        settings_frame = ttk.LabelFrame(main, text="Settings", padding=10)
        settings_frame.grid(row=2, column=0, sticky="ew", pady=(0, 8))
        settings_frame.columnconfigure(1, weight=1)

        # Dataverse URL
        ttk.Label(settings_frame, text="Dataverse URL:").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        self.dataverse_var = tk.StringVar(value=self.settings.dataverse_url)
        ttk.Entry(settings_frame, textvariable=self.dataverse_var).grid(row=0, column=1, columnspan=2, sticky="ew", pady=4)

        # Source picker — radio buttons for SharePoint or Local
        ttk.Label(settings_frame, text="Source:").grid(row=1, column=0, sticky="nw", padx=(0, 8), pady=4)
        source_frame = ttk.Frame(settings_frame)
        source_frame.grid(row=1, column=1, columnspan=2, sticky="ew", pady=4)
        source_frame.columnconfigure(1, weight=1)

        self.source_mode_var = tk.StringVar(
            value="local" if self.settings.local_folder else "sharepoint"
        )

        # Row 1: SharePoint
        ttk.Radiobutton(
            source_frame, text="SharePoint folder URL",
            variable=self.source_mode_var, value="sharepoint",
            command=self._on_source_mode_change,
        ).grid(row=0, column=0, sticky="w")

        self.sharepoint_var = tk.StringVar(value=self.settings.sharepoint_folder_url)
        self.sharepoint_entry = ttk.Entry(source_frame, textvariable=self.sharepoint_var)
        self.sharepoint_entry.grid(row=0, column=1, columnspan=2, sticky="ew", padx=(8, 0))

        # Row 2: Local folder
        ttk.Radiobutton(
            source_frame, text="Local folder",
            variable=self.source_mode_var, value="local",
            command=self._on_source_mode_change,
        ).grid(row=1, column=0, sticky="w", pady=(6, 0))

        self.local_folder_var = tk.StringVar(value=self.settings.local_folder)
        self.local_folder_entry = ttk.Entry(source_frame, textvariable=self.local_folder_var)
        self.local_folder_entry.grid(row=1, column=1, sticky="ew", padx=(8, 4), pady=(6, 0))
        self.local_browse_btn = ttk.Button(
            source_frame, text="Browse...",
            command=self._browse_local_folder,
        )
        self.local_browse_btn.grid(row=1, column=2, sticky="e", pady=(6, 0))

        # Output folder
        ttk.Label(settings_frame, text="Output folder:").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=4)
        self.output_var = tk.StringVar(value=self.settings.output_dir)
        ttk.Entry(settings_frame, textvariable=self.output_var).grid(row=2, column=1, sticky="ew", pady=4)
        ttk.Button(
            settings_frame, text="Browse...",
            command=self._browse_output_folder,
        ).grid(row=2, column=2, sticky="e", padx=(4, 0), pady=4)

        # Existing article mode
        ttk.Label(settings_frame, text="If article exists:").grid(row=3, column=0, sticky="w", padx=(0, 8), pady=4)
        self.existing_var = tk.StringVar(value=self.settings.existing_article_mode)
        existing_combo = ttk.Combobox(
            settings_frame, textvariable=self.existing_var,
            values=["skip", "update", "duplicate"],
            state="readonly", width=12,
        )
        existing_combo.grid(row=3, column=1, sticky="w", pady=4)

        # Save settings button
        save_frame = ttk.Frame(settings_frame)
        save_frame.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(8, 0))
        ttk.Button(save_frame, text="💾 Save Settings", command=self._save_settings).pack(side=tk.RIGHT)

        # Apply initial source-mode visibility
        self._on_source_mode_change()

        # ── Action buttons ────────────────────────────────────────────
        actions_frame = ttk.Frame(main)
        actions_frame.grid(row=3, column=0, sticky="ew", pady=(0, 8))

        self.test_btn = ttk.Button(actions_frame, text="🔌 Test Connection", command=self._on_test_connection)
        self.test_btn.pack(side=tk.LEFT, padx=(0, 6))

        self.dryrun_btn = ttk.Button(actions_frame, text="🧪 Dry Run (preview)", command=lambda: self._start_run(dry_run=True))
        self.dryrun_btn.pack(side=tk.LEFT, padx=6)

        self.run_btn = ttk.Button(actions_frame, text="▶ Run (publish to D365)", command=lambda: self._start_run(dry_run=False))
        self.run_btn.pack(side=tk.LEFT, padx=6)

        self.kb_status_btn = ttk.Button(actions_frame, text="📊 KB Status", command=self._on_kb_status)
        self.kb_status_btn.pack(side=tk.LEFT, padx=6)

        self.open_log_btn = ttk.Button(actions_frame, text="📄 Open Log", command=self._open_last_log, state="disabled")
        self.open_log_btn.pack(side=tk.RIGHT, padx=(6, 0))

        self.open_output_btn = ttk.Button(actions_frame, text="📁 Open Output", command=self._open_output_folder)
        self.open_output_btn.pack(side=tk.RIGHT, padx=(6, 0))

        # ── Progress / log area ───────────────────────────────────────
        progress_frame = ttk.LabelFrame(main, text="Progress", padding=10)
        progress_frame.grid(row=4, column=0, sticky="nsew", pady=(0, 8))
        progress_frame.columnconfigure(0, weight=1)
        progress_frame.rowconfigure(1, weight=1)

        # Progress bar (top of progress area)
        bar_frame = ttk.Frame(progress_frame)
        bar_frame.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        bar_frame.columnconfigure(0, weight=1)

        self.progress_bar = ttk.Progressbar(bar_frame, mode="determinate", maximum=100, value=0)
        self.progress_bar.grid(row=0, column=0, sticky="ew")
        self.progress_label_var = tk.StringVar(value="Ready.")
        ttk.Label(bar_frame, textvariable=self.progress_label_var, width=30).grid(row=0, column=1, sticky="e", padx=(8, 0))

        # Live log text area
        log_frame = ttk.Frame(progress_frame)
        log_frame.grid(row=1, column=0, sticky="nsew")
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log_text = tk.Text(
            log_frame, wrap="word", height=12,
            font=("Consolas" if platform.system() == "Windows" else "Menlo", 10),
        )
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scroll.set)

        # Color tags for log entries
        self.log_text.tag_configure("info", foreground="#1f6feb")
        self.log_text.tag_configure("warning", foreground="#9a6700")
        self.log_text.tag_configure("error", foreground="#cf222e")
        self.log_text.tag_configure("success", foreground="#1a7f37")
        self.log_text.tag_configure("muted", foreground="#656d76")

        # ── Status bar ────────────────────────────────────────────────
        self.status_var = tk.StringVar(value="Ready.")
        status_bar = ttk.Label(
            main, textvariable=self.status_var,
            relief="sunken", anchor="w",
        )
        status_bar.grid(row=5, column=0, sticky="ew")

    # ── UI helpers ─────────────────────────────────────────────────────

    def _on_source_mode_change(self):
        """Enable/disable the SharePoint vs Local fields based on selection."""
        mode = self.source_mode_var.get()
        if mode == "sharepoint":
            self.sharepoint_entry.configure(state="normal")
            self.local_folder_entry.configure(state="disabled")
            self.local_browse_btn.configure(state="disabled")
        else:
            self.sharepoint_entry.configure(state="disabled")
            self.local_folder_entry.configure(state="normal")
            self.local_browse_btn.configure(state="normal")

    def _browse_local_folder(self):
        folder = filedialog.askdirectory(title="Pick a folder with Word files")
        if folder:
            self.local_folder_var.set(folder)

    def _browse_output_folder(self):
        folder = filedialog.askdirectory(title="Pick an output folder")
        if folder:
            self.output_var.set(folder)

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
            self._log("Settings saved.\n", "success")
        except Exception as e:
            messagebox.showerror("Could not save settings", str(e))

    def _set_status(self, msg: str):
        self.status_var.set(msg)

    def _log(self, message: str, tag: str = ""):
        """Append a message to the log text area."""
        self.log_text.configure(state="normal")
        if tag:
            self.log_text.insert(tk.END, message, tag)
        else:
            self.log_text.insert(tk.END, message)
        self.log_text.see(tk.END)

    def _clear_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", tk.END)

    def _set_buttons_enabled(self, enabled: bool):
        """Enable or disable action buttons during a run."""
        state = "normal" if enabled else "disabled"
        for btn in (self.test_btn, self.dryrun_btn, self.run_btn,
                    self.kb_status_btn, self.signin_btn, self.signout_btn):
            btn.configure(state=state)

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
        return self.auth

    def _refresh_auth_status(self):
        """Update the auth indicator label."""
        try:
            auth = self._ensure_auth_client()
            user = auth.get_signed_in_user()
            if user:
                self.auth_status_var.set(f"✓ Signed in as {user}  ({auth.method})")
                self.auth_status_label.configure(foreground="#1a7f37")
            else:
                self.auth_status_var.set(f"Not signed in  ({auth.method})")
                self.auth_status_label.configure(foreground="#656d76")
        except RuntimeError as e:
            self.auth_status_var.set(f"⚠ Auth not available: {e}")
            self.auth_status_label.configure(foreground="#cf222e")

    def _on_signin(self):
        """Trigger interactive sign-in by acquiring a Graph token."""
        self._set_status("Signing in — check your browser...")
        self._log("Signing in...\n", "info")

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
        if not messagebox.askyesno("Sign out", "Are you sure you want to sign out?"):
            return
        try:
            auth = self._ensure_auth_client()
            auth.sign_out()
            self.auth = None  # force re-init on next operation
            self._refresh_auth_status()
            self._log("Signed out.\n", "info")
            self._set_status("Signed out.")
        except Exception as e:
            messagebox.showerror("Sign out failed", str(e))

    # ── Test connection ───────────────────────────────────────────────

    def _on_test_connection(self):
        s = self._collect_settings()
        ok, error = s.is_valid_for_run()
        if not ok:
            messagebox.showerror("Settings incomplete", error)
            return

        self._clear_log()
        self._log("Testing connection...\n", "info")
        self._set_status("Testing connection...")
        self._set_buttons_enabled(False)

        def worker():
            try:
                # Update auth client with current settings
                self.settings = s
                self.auth = None
                auth = self._ensure_auth_client()

                # Test Dataverse access
                self.event_queue.put(("log", "  Acquiring Dataverse token...\n", "muted"))
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

                # Test SharePoint if configured
                if s.input_mode == "sharepoint":
                    self.event_queue.put(("log", "  Acquiring Graph token...\n", "muted"))
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
                self.event_queue.put(("test_done", f"err:{e}"))

        threading.Thread(target=worker, daemon=True).start()

    # ── KB status ─────────────────────────────────────────────────────

    def _on_kb_status(self):
        s = self._collect_settings()
        if not s.dataverse_url:
            messagebox.showerror("Dataverse URL required", "Please enter the Dataverse URL.")
            return
        self._set_buttons_enabled(False)
        self._log("Fetching KB article counts...\n", "info")

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
                lines = ["\n  KB Article Summary\n", "  " + "=" * 35 + "\n"]
                for status in sorted(counts.keys() - {"Total"}):
                    lines.append(f"  {status:<20} {counts[status]:>6}\n")
                lines.append("  " + "-" * 35 + "\n")
                lines.append(f"  {'Total':<20} {counts.get('Total', 0):>6}\n")
                lines.append("  " + "=" * 35 + "\n\n")
                self.event_queue.put(("log_block", "".join(lines), "info"))
                self.event_queue.put(("status_done", "Ready."))
            except Exception as e:
                self.event_queue.put(("test_done", f"err:{e}"))

        threading.Thread(target=worker, daemon=True).start()

    # ── Run / Dry run ─────────────────────────────────────────────────

    def _start_run(self, dry_run: bool):
        s = self._collect_settings()
        ok, error = s.is_valid_for_run()
        if not ok:
            messagebox.showerror("Settings incomplete", error)
            return

        # Confirm live run
        if not dry_run:
            confirm = messagebox.askyesno(
                "Confirm run",
                "This will publish articles to Dynamics 365.\n\n"
                f"Existing article mode: {s.existing_article_mode}\n\n"
                "Continue?",
            )
            if not confirm:
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
        self.progress_label_var.set("Starting...")
        self._set_buttons_enabled(False)
        self.last_log_path = None
        self.last_run_log_path = None

        mode_label = "DRY RUN" if dry_run else "LIVE RUN"
        self._log(f"=== {mode_label} ===\n", "info")
        self._set_status(f"{mode_label} in progress...")

        def worker():
            try:
                self.auth = None
                auth = self._ensure_auth_client()
                config = LoadConfig.from_settings(s, dry_run=dry_run)

                # Pre-flight auth (live mode only)
                if not dry_run:
                    self.event_queue.put(("log", "Authenticating...\n", "info"))
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
                self.event_queue.put(("run_failed", str(e)))

        self.worker_thread = threading.Thread(target=worker, daemon=True)
        self.worker_thread.start()

    # ── Open helpers ──────────────────────────────────────────────────

    def _open_last_log(self):
        if self.last_log_path and Path(self.last_log_path).exists():
            _open_path_in_explorer(self.last_log_path)
        else:
            messagebox.showinfo("No log yet", "Run a load first to generate a log file.")

    def _open_output_folder(self):
        s = self._collect_settings()
        out = Path(s.output_dir or "./output")
        out.mkdir(parents=True, exist_ok=True)
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
        # Re-arm
        self.root.after(100, self._drain_event_queue)

    def _handle_event(self, event: tuple):
        kind = event[0]
        if kind == "log":
            _, msg, tag = event
            self._log(msg, tag)
        elif kind == "log_block":
            _, msg, tag = event
            self._log(msg, tag)
        elif kind == "auth_ok":
            user = event[1]
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
            self._log(f"⚠ Sign-in failed: {err}\n", "error")
            self._set_status("Sign-in failed.")
            messagebox.showerror("Sign-in failed", err)
            self._set_buttons_enabled(True)
        elif kind == "test_done":
            outcome = event[1]
            if outcome == "ok":
                self._log("Connection test passed.\n", "success")
                self._set_status("Connection OK.")
            else:
                err = outcome[4:]
                self._log(f"⚠ {err}\n", "error")
                self._set_status("Connection failed.")
                messagebox.showerror("Connection test failed", err)
            self._set_buttons_enabled(True)
            self._refresh_auth_status()
        elif kind == "status_done":
            self._set_status(event[1])
            self._set_buttons_enabled(True)
            self._refresh_auth_status()
        elif kind == "progress_event":
            self._render_progress_event(event[1])
        elif kind == "run_done":
            result = event[1]
            self._on_run_complete(result)
        elif kind == "run_failed":
            err = event[1]
            self._log(f"\n⚠ Run failed: {err}\n", "error")
            self._set_status("Run failed.")
            messagebox.showerror("Run failed", err)
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
                self.progress_label_var.set(f"{event.current}/{event.total}")
        elif event.kind == "file_done":
            status_str = ", ".join(f"{k}: {v}" for k, v in event.status.items())
            tag = "error" if any("ERROR" in str(v) for v in event.status.values()) else ""
            self._log(f"  [{event.current}/{event.total}] {event.file_name}  →  {status_str}\n", tag)
            if event.total > 0:
                pct = (event.current / event.total) * 100
                self.progress_bar.configure(value=pct, maximum=100)
                self.progress_label_var.set(f"{event.current}/{event.total}")
        elif event.kind == "summary":
            self._log(f"\n{event.message}\n", "info")

    def _on_run_complete(self, result):
        self.progress_bar.configure(value=100)
        self.progress_label_var.set("Done.")
        self.last_log_path = result.log_path
        self.last_run_log_path = result.run_log_path
        self._log("\n" + "=" * 60 + "\n", "muted")
        self._log("Run complete!\n", "success" if result.success else "warning")
        self._log(
            f"  Converted: {result.converted}, Created: {result.created}, "
            f"Updated: {result.updated}, Skipped: {result.skipped}, Errors: {result.errors}\n",
            "info",
        )
        if result.log_path:
            self._log(f"  Detail log: {result.log_path}\n", "muted")
        if result.run_log_path:
            self._log(f"  Run log:    {result.run_log_path}\n", "muted")
        self._log("=" * 60 + "\n", "muted")

        self.open_log_btn.configure(state="normal")
        self._set_buttons_enabled(True)
        if result.errors:
            self._set_status(f"Run finished with {result.errors} error(s).")
        else:
            self._set_status("Run complete.")
        self._refresh_auth_status()


def launch_gui():
    """Launch the Tkinter GUI. Blocks until the window is closed."""
    # Ensure logging is configured (the service.py setup_file_logging adds a file handler per run)
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.DEBUG)

    root = tk.Tk()
    # Try to use a more modern theme on Windows
    try:
        style = ttk.Style(root)
        if "vista" in style.theme_names() and platform.system() == "Windows":
            style.theme_use("vista")
        elif "aqua" in style.theme_names() and platform.system() == "Darwin":
            style.theme_use("aqua")
    except Exception:
        pass

    KBLoaderGUI(root)
    root.mainloop()


if __name__ == "__main__":
    launch_gui()
