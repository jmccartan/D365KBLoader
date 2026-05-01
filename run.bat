@echo off
REM =====================================================================
REM D365 Knowledge Base Loader
REM Copyright (c) 2026 John McCartan
REM Licensed under the MIT License. See the LICENSE file in the project
REM root for the full text.
REM =====================================================================
REM ============================================================
REM  D365 Knowledge Base Loader — Windows launcher
REM  Double-click this file to run the app.
REM
REM  On first run, it creates a Python virtual environment in
REM  .venv\ and installs dependencies. Subsequent runs are fast.
REM ============================================================

setlocal
cd /d "%~dp0"

REM Look for Python on PATH
where python >nul 2>nul
if errorlevel 1 (
  where py >nul 2>nul
  if errorlevel 1 (
    echo.
    echo  ERROR: Python is not installed.
    echo.
    echo  Install Python 3.10 or newer from:
    echo    https://www.python.org/downloads/
    echo.
    echo  IMPORTANT: During install, check the box
    echo  "Add python.exe to PATH".
    echo.
    pause
    exit /b 1
  )
  set PY=py -3
) else (
  set PY=python
)

REM Create venv if missing
if not exist ".venv\Scripts\python.exe" (
  echo.
  echo  First-run setup — creating Python environment...
  echo.
  %PY% -m venv .venv
  if errorlevel 1 (
    echo  Failed to create virtual environment. See the error above.
    pause
    exit /b 1
  )
  call .venv\Scripts\activate.bat
  echo  Upgrading pip...
  python -m pip install --upgrade pip setuptools wheel
  echo  Installing dependencies (using prebuilt wheels)...
  pip install --prefer-binary -r requirements.txt
  if errorlevel 1 (
    echo.
    echo  Failed to install dependencies. See the error above.
    pause
    exit /b 1
  )
) else (
  call .venv\Scripts\activate.bat
)

REM Launch the GUI
python -m kb_loader
if errorlevel 1 (
  echo.
  echo  The app exited with an error.
  pause
)
endlocal
