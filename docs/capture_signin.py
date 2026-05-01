# =====================================================================
# D365 Knowledge Base Loader
# Copyright (c) 2026 John McCartan
# Licensed under the MIT License. See the LICENSE file in the project
# root for the full text.
# =====================================================================

"""Capture a screenshot of the device code dialog overlay (sanitized)."""
import time
from pathlib import Path

import ttkbootstrap as ttk
from PIL import ImageGrab

from kb_loader import gui as gui_module
from kb_loader.gui import KBLoaderGUI, DEFAULT_THEME, _enable_windows_dpi_awareness

# Suppress browser auto-open during capture
gui_module._open_url_in_browser = lambda url: True

_enable_windows_dpi_awareness()
root = ttk.Window(themename=DEFAULT_THEME)
gui = KBLoaderGUI(root)

# Sanitize main window so it doesn't leak real values
gui.dataverse_var.set("https://contoso.crm.dynamics.com/")
gui.sharepoint_var.set("https://contoso.sharepoint.com/sites/KnowledgeBase/Shared Documents/KB Articles")
gui.auth_status_var.set("Signed in: kb-admin@contoso.onmicrosoft.com")
gui.auth_method_var.set("via msal")

gui._log("Starting sign-in…\n", "info")
root.update_idletasks()
root.update()
time.sleep(0.3)

# Show device code dialog with placeholder code
gui._show_device_code_dialog("F7Q3K9MAJ", "https://microsoft.com/devicelogin")
root.update_idletasks()
root.update()
time.sleep(0.5)
root.update()

x = root.winfo_rootx()
y = root.winfo_rooty()
w = root.winfo_width()
h = root.winfo_height()
img = ImageGrab.grab(bbox=(x - 2, y - 32, x + w + 2, y + h + 2))

out_path = Path("docs/screenshot_signin.png")
img.save(out_path)
print(f"Saved {out_path}  ({img.width}x{img.height})")

root.destroy()
