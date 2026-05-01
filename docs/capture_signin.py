"""Capture a screenshot of the device code dialog overlay."""
import time
from pathlib import Path

import ttkbootstrap as ttk
from PIL import ImageGrab

from kb_loader.gui import KBLoaderGUI, DEFAULT_THEME, _enable_windows_dpi_awareness

_enable_windows_dpi_awareness()
root = ttk.Window(themename=DEFAULT_THEME)
gui = KBLoaderGUI(root)
gui._log("Starting sign-in…\n", "info")
root.update_idletasks()
root.update()
time.sleep(0.3)

# Show device code dialog
gui._show_device_code_dialog("F7Q3K9MAJ", "https://microsoft.com/devicelogin")
root.update_idletasks()
root.update()
time.sleep(0.5)
root.update()

# Capture the area covering both windows
x = root.winfo_rootx()
y = root.winfo_rooty()
w = root.winfo_width()
h = root.winfo_height()
img = ImageGrab.grab(bbox=(x - 2, y - 32, x + w + 2, y + h + 2))

out_path = Path("docs/screenshot_signin.png")
img.save(out_path)
print(f"Saved {out_path}  ({img.width}x{img.height})")

root.destroy()
