"""Capture the sharing-link recovery dialog (sanitized)."""
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

gui.dataverse_var.set("https://contoso.crm.dynamics.com/")
gui.auth_status_var.set("Signed in: kb-admin@contoso.onmicrosoft.com")
gui.auth_method_var.set("via msal")

root.update_idletasks()
root.update()
time.sleep(0.3)

gui._show_sharing_link_recovery_dialog(
    "https://contoso.sharepoint.com/:f:/s/KnowledgeBase/abc123token456",
    "Test",
)
root.update_idletasks()
root.update()
time.sleep(0.5)
root.update()

# Find the dialog window and capture it
for w in root.winfo_children():
    if isinstance(w, ttk.Toplevel) or w.winfo_class() == "Toplevel":
        x = w.winfo_rootx()
        y = w.winfo_rooty()
        ww = w.winfo_width()
        wh = w.winfo_height()
        img = ImageGrab.grab(bbox=(x - 2, y - 32, x + ww + 2, y + wh + 2))
        img.save("docs/screenshot_recovery.png")
        print(f"Saved docs/screenshot_recovery.png ({img.width}x{img.height})")
        break

root.destroy()
