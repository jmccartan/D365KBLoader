# D365 Knowledge Base Loader

A simple cross-platform tool that bulk-loads Word documents (`.docx` and `.doc`) into Dynamics 365 Knowledge Base articles. Works on **Windows** and **macOS**.

![Screenshot of the D365 Knowledge Base Loader GUI](docs/screenshot.png)

---

## Quick Start

### 1. Install Python (one-time)

Download Python **3.10 or newer** from [python.org/downloads](https://www.python.org/downloads/).

> **Windows:** During install, check the box that says **"Add python.exe to PATH"**.
> **macOS:** Or use Homebrew: `brew install python`.

### 2. Run the app

Double-click the launcher for your operating system:

| OS | Launcher | Notes |
|----|----------|-------|
| **Windows** | `run.bat` | Double-click in File Explorer |
| **macOS** | `run.command` | If macOS blocks it the first time, **right-click → Open** |

The first run takes about a minute to set up the Python environment and install dependencies. After that the app opens in a few seconds.

### 3. Use the app

1. Click **Sign In** — a sign-in code appears in a dialog and your browser opens to https://microsoft.com/devicelogin. Paste the code, sign in.
2. Fill in the **Dataverse URL** (e.g. `https://your-org.crm.dynamics.com`)
3. Pick your **source** (SharePoint URL or local folder)
4. Click **Test Connection** to verify everything works
5. Click **Dry Run** first to preview the conversion
6. Click **▶ Run** to publish to D365

That's it. Your settings are remembered for next time.

---

## What each section does

### 1. Account

The colored dot tells you at a glance whether you're signed in:

| Dot | Meaning |
|-----|---------|
| 🟢 Green | Signed in — ready to use |
| ⚪ Grey  | Not signed in yet |
| 🔴 Red   | Authentication isn't available (Azure CLI not installed) |

The label below shows **how** you're authenticated. The app uses **hybrid auth** by default:
- **Microsoft Graph** (SharePoint enumeration) → MSAL device-code sign-in
- **Dataverse** (article CRUD) → your existing Azure CLI session

Tokens are cached in `~/.d365kbloader/` so future runs don't re-prompt.

### 2. Settings

| Field | Purpose |
|-------|---------|
| **Dataverse URL** | The URL of your D365 environment (e.g. `https://your-org.crm.dynamics.com`) |
| **Source: SharePoint folder URL** | Read documents from SharePoint via Microsoft Graph. Paste **either** a URL from your browser address bar **or** a sharing link (`Copy link` from SharePoint). Click the **?** button next to the field for visual help. |
| **Source: Local folder** | Read documents from a folder on your computer (e.g. an OneDrive-synced SharePoint folder) |
| **Output folder** | Where HTML files, the detail log, and the Excel run log are written. Defaults to `./output` |
| **If article exists** | What to do when a Knowledge Article with the same title already exists: |
| | • **skip** — leave the existing one alone (safest, default) |
| | • **update** — overwrite the existing article's content |
| | • **duplicate** — create a new article anyway |

The **Run** button auto-saves your settings before each run, so you rarely need to click **Save settings** manually. Settings persist to `~/.d365kbloader/settings.json` and survive reinstalls.

### 3. Action buttons

| Button | What it does |
|--------|-------------|
| **Test Connection** | Verifies your Dataverse URL works and the source (SharePoint / local folder) is reachable. Reports the number of files found and total KB articles |
| **KB Status** | Shows how many Knowledge Articles exist by status (Draft, Published, Archived, etc.) |
| **Dry Run** | Converts files to HTML and writes the run log, but **does NOT publish** to D365. Useful for previewing |
| **▶ Run** | The real thing — converts, publishes to D365, and updates Knowledge Articles |

> **Recommendation:** Always do a **Dry Run** first to verify file conversion and styling, then **Run** when you're happy with the preview.

### 4. Progress

While a run is in progress:

- **Progress bar** fills as files are processed
- **`X / Y  (NN%)`** counter on the right shows current file number and percent complete
- **Live output** displays each file's outcome on a separate line:
  - `content: yes/EMPTY` — whether the document had content
  - `html: saved/skipped` — whether an HTML file was written (skipped for empty docs)
  - `kb: created/updated/skipped (exists)/dry run/ERROR` — what happened in D365

Above the log:
- **Clear** — wipe the live output area
- **Open output folder** — open the folder with HTML files & logs in Explorer/Finder
- **Open detail log** — open the timestamped `.log` file with full debug detail (enabled after a run)

---

## What gets created in D365

Each Word document becomes a Knowledge Article with these fields:

| Field | Value |
|-------|-------|
| **Title** | Filename without extension (e.g. `Refund Policy.docx` → `Refund Policy`) |
| **Content** | HTML converted from the Word document, with **inline styles applied** (see below) |
| **Language** | English (locale 1033) |
| **Creation mode** | Manual |
| **Description** | `Auto-imported from SharePoint: <source path>` |
| **Keywords** | The source file path (for traceability) |
| **Status** | Published (transitioned through Draft → Approved → Published) |

Empty documents (no content after conversion) are **skipped** — no HTML file is written and no KB article is created.

### Inline styling

D365's rich-text editor strips `<style>` blocks and external stylesheets, so all formatting must live inline on each element. The app post-processes mammoth's HTML output and applies a Microsoft Fluent-inspired theme:

- **Headings** — `<h1>` and `<h2>` get a Fluent blue accent; sizes scale down through `<h6>`
- **Body** — Segoe UI 14px, line-height 1.5, max-width 820px
- **Links** — Fluent blue, no underline
- **Lists** — clean indentation and spacing
- **Code** — monospace, light gray pill background; code blocks get a blue left-border accent
- **Tables** — bordered cells with a shaded header row

Plus **smart callouts**: paragraphs starting with these prefixes are auto-converted to colored callout boxes:

| Prefix | Style |
|--------|-------|
| `Note:` | Yellow with amber border |
| `Tip:` / `Hint:` | Green with green border |
| `Warning:` / `Important:` / `Caution:` | Red with red border |

---

## Authentication

The app uses **hybrid authentication** by default — this works for any user in any tenant **without any app registration**:

| API | Method | Why |
|-----|--------|-----|
| **Microsoft Graph** (SharePoint) | MSAL with the Microsoft Graph PowerShell public client ID | This well-known Microsoft client is pre-authorized for `Sites.Read.All` and `Files.Read.All`, which are needed to enumerate SharePoint document libraries. Azure CLI's Graph token doesn't have these scopes. |
| **Dataverse** (D365) | Azure CLI (`az login`) | Works reliably, no app registration needed, your existing `az` session is reused. |

When you click **Sign In** in the GUI, a dialog appears showing a **device-code** like `F7Q3K9MAJ`:

![Sign-in dialog](docs/screenshot_signin.png)

The browser opens automatically to https://microsoft.com/devicelogin — paste the code there and sign in with your Microsoft account. The dialog closes automatically once sign-in completes.

> **Why a code instead of a popup?** Browser-based sign-in windows often get hidden behind other apps. The device-code flow is more reliable because the code is always visible right in the app.

### Tenant policy override

If your tenant blocks the default Microsoft Graph PowerShell client, you have two options:

1. **Register your own Entra app** with delegated `Sites.Read.All` and `Files.Read.All` permissions, then set `AZURE_CLIENT_ID=<your-app-id>` in `.env`.
2. **Force az CLI mode** by setting `KB_LOADER_AUTH=az_cli` in `.env`. (SharePoint enumeration may have limited functionality in this mode — see Troubleshooting.)

### Required permissions

Your Microsoft account needs:
- **D365**: any role that lets you create/edit Knowledge Articles (e.g. Knowledge Manager)
- **SharePoint** *(if using SharePoint source)*: **direct membership** of the site (a sharing link alone is not enough — the user must be added to Members or have direct access)

---

## What goes in the output folder

Everything writes to the **single output folder** you choose (default: `./output`):

```
output/
├── 📁 KB Main Folder/                                ← matches your source structure
│   ├── 📁 Dynamic Rebooking Tool/
│   │   ├── 📄 Customer Accepts Auto Reaccommodated Flight….html
│   │   └── 📄 Customer Rebooks to Finnair….html
│   └── 📄 Refund Policy.html
├── 📄 kb_loader_20260501_080100.log                  ← detail log (per run)
└── 📄 kb_loader_log_20260501_080100.xlsx             ← Excel run log (per run)
```

- **HTML files** — one per Word document, in subfolders matching the source folder structure. Empty documents are skipped (no HTML written). Each HTML file is fully styled and ready to paste into D365.
- **`kb_loader_YYYYMMDD_HHMMSS.log`** — full timestamped log with every step (great for troubleshooting)
- **`kb_loader_log_YYYYMMDD_HHMMSS.xlsx`** — Excel summary with one row per file:

| Column | Description |
|--------|-------------|
| File Name | Original filename |
| Folder Path | Subfolder relative to source root |
| File Size | Bytes |
| Has Content | Yes / No |
| HTML Saved | Yes / No |
| Published to KB | Yes / No / Skipped |
| KB Action | Created / Updated / Skipped / Dry Run / Error |
| Article ID | Dataverse `knowledgearticleid` |
| Error | Error message if processing failed |

Cells are color-coded (green = yes, red = no, yellow = skipped). The Excel file also includes a **before/after KB article count comparison** at the top.

> Click **Open output folder** in the GUI to jump straight to it. The live output area also tells you how many HTML files / logs are currently there.

---

## Troubleshooting

| Symptom | Try this |
|---------|----------|
| **"Python is not installed"** when launching | Install Python from [python.org](https://www.python.org/downloads/). On Windows, check "Add to PATH" during install |
| **macOS: "cannot be opened because the developer cannot be verified"** | Right-click `run.command` → **Open** → click **Open** in the dialog |
| **Sign-in keeps re-prompting** | Click **Sign Out** in the app, then **Sign In** again |
| **"The SharePoint site was found, but no document libraries are visible"** | Your account isn't a direct Member of the site — only the sharing link gives you access. Either ask the site owner to add you as a Member, OR pick a SharePoint folder you're already a member of, OR sync via OneDrive and use Local folder mode |
| **SharePoint sharing-link recovery dialog opens** | Sign-in works but Graph couldn't auto-resolve the sharing link. Click **Open link in browser**, copy the URL from the address bar after the page loads, and paste it back into the dialog's Step 3 field |
| **Legacy `.doc` files don't convert** | Install [LibreOffice](https://www.libreoffice.org/download/). Modern `.docx` files don't need it |
| **GUI looks tiny on a HiDPI monitor** | The app enables Windows DPI awareness automatically. If still tiny, try moving the window to your primary display |
| **First-run dependency install fails on ARM64 Windows** | The launcher uses `--prefer-binary` to grab prebuilt wheels. If it still fails, run `python -m pip install --upgrade pip setuptools wheel` then retry the launcher |

For deeper investigation, click **Open detail log** in the GUI (or open the timestamped `.log` file in your output folder) to see the full trace.

---

## Command-line usage (advanced)

The CLI is still available for scripted/automated runs:

```bash
# Launch the GUI (default with no args)
python -m kb_loader

# CLI mode
python -m kb_loader --local-folder "C:\docs" --dry-run
python -m kb_loader --sharepoint-url "https://..." --existing update
python -m kb_loader --kb-status
python -m kb_loader --help
```

CLI args override settings from `.env` and `~/.d365kbloader/settings.json`.

---

## Cross-platform support

- ✅ **Windows 10/11** (x64 and ARM64) — primary development platform
- ✅ **macOS 12+** — uses native Aqua look-and-feel via Tk
- ✅ **Linux** — should work but not actively tested

Platform-specific behavior:
- **Fonts** — Segoe UI on Windows, SF Pro on macOS, DejaVu on Linux
- **HiDPI** — Windows DPI awareness is enabled automatically; macOS Retina is native
- **File / browser opening** — `os.startfile` on Windows, `open` on macOS, `xdg-open` on Linux
- **Azure CLI calls** — uses `shell=True` only on Windows for `az.cmd` resolution
- **Line endings** — `.gitattributes` enforces LF on `*.command` so Mac launchers work after a Windows clone

---

## Project structure

```
D365KBLoader/
├── run.bat                    ← Windows launcher (double-click)
├── run.command                ← macOS launcher (double-click)
├── requirements.txt           ← Python dependencies
├── .env.example               ← Optional environment overrides
├── .gitattributes             ← Line-ending rules for cross-platform
├── README.md                  ← this file
├── docs/
│   ├── screenshot.png         ← Main UI screenshot
│   ├── screenshot_signin.png  ← Sign-in dialog screenshot
│   ├── screenshot_recovery.png ← Sharing-link recovery dialog
│   ├── capture_screenshot.py  ← Utility to regenerate the main screenshot
│   ├── capture_signin.py      ← Utility to regenerate the sign-in screenshot
│   └── capture_recovery.py    ← Utility to regenerate the recovery dialog screenshot
└── kb_loader/
    ├── __init__.py
    ├── __main__.py            ← CLI / GUI entry point (GUI is default)
    ├── gui.py                 ← Tkinter + ttkbootstrap GUI
    ├── service.py             ← Core load logic (used by both CLI and GUI)
    ├── settings.py            ← User-profile settings store
    ├── auth.py                ← Hybrid MSAL + Azure CLI authentication
    ├── sharepoint_client.py   ← SharePoint via Graph API
    ├── dataverse_client.py    ← D365 Knowledge Article CRUD
    ├── converter.py           ← Word → HTML conversion (mammoth + LibreOffice)
    ├── styles.py              ← Inline-style post-processor for D365
    ├── config.py              ← Legacy CLI Config dataclass (used by clients)
    └── run_log.py             ← Excel run log generator
```
