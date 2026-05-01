"""Capture the recovery dialog for visual review."""
import time
from pathlib import Path

import ttkbootstrap as ttk
from PIL import ImageGrab

# No-op browser
import webbrowser
webbrowser.open = lambda url: None

from kb_loader.gui import KBLoaderGUI, DEFAULT_THEME, _enable_windows_dpi_awareness

_enable_windows_dpi_awareness()
root = ttk.Window(themename=DEFAULT_THEME)
gui = KBLoaderGUI(root)
root.update_idletasks()
root.update()
time.sleep(0.3)

gui._show_sharing_link_recovery_dialog(
    "https://absx29592600.sharepoint.com/:f:/s/McDemo/IgCoz3LeG3D4TKs_GVYgebhhAQBrNsGNoiftdoDXAJ2DfEs",
    "Test",
)
root.update_idletasks()
root.update()
time.sleep(0.5)
root.update()

# Capture the entire screen area for the dialog
# Find the dialog window
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
