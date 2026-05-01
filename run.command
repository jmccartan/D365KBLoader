#!/bin/bash
# =====================================================================
# D365 Knowledge Base Loader
# Copyright (c) 2026 John McCartan
# Licensed under the MIT License. See the LICENSE file in the project
# root for the full text.
# =====================================================================
# ============================================================
#  D365 Knowledge Base Loader — Mac/Linux launcher
#  Double-click this file (Mac) or run from terminal (Linux).
#
#  On first run, it creates a Python virtual environment in
#  .venv/ and installs dependencies. Subsequent runs are fast.
# ============================================================

set -e
cd "$(dirname "$0")"

# Find a Python interpreter (3.10+)
if command -v python3 >/dev/null 2>&1; then
  PY=python3
elif command -v python >/dev/null 2>&1; then
  PY=python
else
  echo
  echo "  ERROR: Python is not installed."
  echo
  echo "  Install Python 3.10 or newer:"
  echo "    Mac:   brew install python"
  echo "           or https://www.python.org/downloads/"
  echo "    Linux: use your package manager"
  echo
  read -n 1 -s -r -p "Press any key to close..."
  exit 1
fi

# Create venv if missing
if [ ! -x ".venv/bin/python" ]; then
  echo
  echo "  First-run setup — creating Python environment..."
  echo
  "$PY" -m venv .venv
  source .venv/bin/activate
  echo "  Upgrading pip..."
  python -m pip install --upgrade pip setuptools wheel
  echo "  Installing dependencies (using prebuilt wheels)..."
  pip install --prefer-binary -r requirements.txt
else
  source .venv/bin/activate
fi

# Launch the GUI
python -m kb_loader
EXIT_CODE=$?

if [ $EXIT_CODE -ne 0 ]; then
  echo
  echo "  The app exited with an error."
  read -n 1 -s -r -p "Press any key to close..."
fi
