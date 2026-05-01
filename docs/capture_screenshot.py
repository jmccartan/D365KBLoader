"""Capture a screenshot of the GUI for the README.

Launches the GUI, waits for layout to settle, takes a screenshot of the window
region, then closes. Saves to docs/screenshot.png.
"""
import sys
import time
from pathlib import Path

import ttkbootstrap as ttk
from PIL import ImageGrab

from kb_loader.gui import KBLoaderGUI, DEFAULT_THEME, _enable_windows_dpi_awareness

_enable_windows_dpi_awareness()
root = ttk.Window(themename=DEFAULT_THEME)
gui = KBLoaderGUI(root)

# Add some sample log content so the screenshot looks more useful
gui._log("✓ Settings loaded.\n", "success")
gui._log("Ready to test connection or run a load.\n", "info")
gui._log("\n[Hint] Click 'Test Connection' first to verify everything works.\n", "muted")

# Make sure layout is settled
root.update_idletasks()
root.update()
time.sleep(0.5)
root.update()

# Get window bounds
x = root.winfo_rootx()
y = root.winfo_rooty()
w = root.winfo_width()
h = root.winfo_height()

# Capture some space around the title bar too
title_bar_height = 32
img = ImageGrab.grab(bbox=(x - 2, y - title_bar_height, x + w + 2, y + h + 2))

out_path = Path("docs/screenshot.png")
out_path.parent.mkdir(parents=True, exist_ok=True)
img.save(out_path)
print(f"Saved {out_path}  ({img.width}x{img.height})")

root.destroy()
