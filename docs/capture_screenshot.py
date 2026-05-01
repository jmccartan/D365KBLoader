"""Capture a screenshot of the GUI for the README, with sanitized demo values.

Launches the GUI with placeholder URLs/email so the screenshot doesn't leak
real tenant info, takes a screenshot, then closes. Saves to docs/screenshot.png.
"""
import time
from pathlib import Path

import ttkbootstrap as ttk
from PIL import ImageGrab

from kb_loader.gui import KBLoaderGUI, DEFAULT_THEME, _enable_windows_dpi_awareness

_enable_windows_dpi_awareness()
root = ttk.Window(themename=DEFAULT_THEME)
gui = KBLoaderGUI(root)

# Sanitize displayed values so the screenshot doesn't leak real tenant info
gui.dataverse_var.set("https://contoso.crm.dynamics.com/")
gui.sharepoint_var.set("https://contoso.sharepoint.com/sites/KnowledgeBase/Shared Documents/KB Articles")
gui.output_var.set("./output")
gui.source_mode_var.set("sharepoint")
gui._on_source_mode_change()

# Force a clean signed-in indicator with a placeholder identity
gui.auth_icon_var.set("●")
gui.auth_icon.configure(bootstyle="success")
gui.auth_status_var.set("Signed in: kb-admin@contoso.onmicrosoft.com")
gui.auth_method_var.set("via msal")

# Sample live log
gui._log("✓ Settings loaded.\n", "success")
gui._log("Ready to test connection or run a load.\n", "info")
gui._log("\n[Hint] Click 'Test Connection' first to verify everything works.\n", "muted")

# Make sure layout is settled
root.update_idletasks()
root.update()
time.sleep(0.5)
root.update()

x = root.winfo_rootx()
y = root.winfo_rooty()
w = root.winfo_width()
h = root.winfo_height()

title_bar_height = 32
img = ImageGrab.grab(bbox=(x - 2, y - title_bar_height, x + w + 2, y + h + 2))

out_path = Path("docs/screenshot.png")
out_path.parent.mkdir(parents=True, exist_ok=True)
img.save(out_path)
print(f"Saved {out_path}  ({img.width}x{img.height})")

root.destroy()
