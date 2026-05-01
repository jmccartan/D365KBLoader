# D365 Knowledge Base Loader

A simple cross-platform tool that bulk-loads Word documents (`.docx` and `.doc`) into Dynamics 365 Knowledge Base articles.

## Quick Start

### 1. Install Python (one-time)

Download Python 3.10 or newer from [python.org/downloads](https://www.python.org/downloads/).

> **Windows users:** During install, check the box that says **"Add python.exe to PATH"**.

### 2. Run the app

Just double-click the launcher for your operating system:

- **Windows:** double-click `run.bat`
- **Mac:** double-click `run.command` *(if Mac blocks it the first time, right-click → Open)*

The first run takes a minute to set things up. After that it opens immediately.

### 3. In the app

1. Click **Sign In** (a browser window opens, sign in with your Microsoft account)
2. Fill in the **Dataverse URL** (e.g. `https://your-org.crm.dynamics.com`)
3. Pick your **source**:
   - **SharePoint folder URL** — copied from your browser's address bar, OR
   - **Local folder** — click Browse to pick a folder on your computer
4. Click **🔌 Test Connection** to verify everything works
5. Click **🧪 Dry Run** first to preview the conversion (nothing is published)
6. Click **▶ Run** to publish to D365

The app keeps a detailed log of every run, plus an Excel summary you can review.

---

## What the app does

1. Finds every `.docx` and `.doc` file in your folder (including subfolders)
2. Converts each one to clean HTML
3. Saves a local HTML copy for reference
4. Creates a Knowledge Article in D365 with:
   - **Title** = the filename (without extension)
   - **Content** = the converted HTML
   - **Status** = Published
   - **Language** = English

If a file is empty (no content), it's skipped — no HTML or KB article is created.

---

## Authentication

The app uses your existing Microsoft login. There are two ways it can authenticate:

| Method | When it's used |
|--------|----------------|
| **Azure CLI** | If you've run `az login` on your machine. The app picks up your existing session — no setup needed. |
| **MSAL** | If your IT admin has registered an Entra app and provided an `AZURE_CLIENT_ID`. Better for users without an Azure subscription. |

In both cases, the app handles tenant detection automatically — **you don't need to enter a tenant ID**.

### Required permissions

Your Microsoft account needs:
- Access to the D365 environment (any role that lets you create knowledge articles)
- For SharePoint mode: read access to the SharePoint site

---

## Settings

The app saves your preferences (Dataverse URL, source folder, etc.) to `~/.d365kbloader/settings.json` so they're remembered across runs and even if you reinstall.

You can also use a `.env` file in the project folder to override settings — useful for advanced users or scripted runs.

---

## Idempotency (re-running)

When you re-run, the app checks each Word file's title against existing articles:

| Mode | What happens to existing articles |
|------|----------------------------------|
| **skip** *(default)* | Leave them alone, don't recreate |
| **update** | Overwrite the content with the new version |
| **duplicate** | Create a new article anyway |

---

## Run logs

After every run, two files are written to your output folder:

- **`kb_loader_YYYYMMDD_HHMMSS.log`** — detailed log of every step (helpful for troubleshooting)
- **`kb_loader_log_YYYYMMDD_HHMMSS.xlsx`** — Excel summary with one row per file (easy to share with colleagues)

The Excel file shows file names, folder paths, content status, KB action (Created/Updated/Skipped/Error), and article IDs.

---

## Troubleshooting

| Symptom | What to try |
|---------|-------------|
| **"Python is not installed"** when launching | Install Python from python.org and check "Add to PATH" during install |
| **Sign-in dialog appears repeatedly** | Click Sign Out, then Sign In again. If using Azure CLI, try `az logout && az login` in a terminal |
| **"Could not get a token from Azure CLI"** | Your account may not have access to that Dataverse environment. Contact your D365 admin |
| **SharePoint URL not working** | Use the URL straight from your browser's address bar — sharing links (the kind from "Copy link") aren't supported |
| **Legacy `.doc` files not converting** | Install [LibreOffice](https://www.libreoffice.org/download/). Modern `.docx` files don't need it |

For deeper investigation, open the **detail log** (timestamped `.log` file in your output folder) to see the full trace of every step.

---

## Command-line usage (advanced)

You can also run from a terminal:

```bash
# Launch the GUI
python -m kb_loader

# Or use the CLI
python -m kb_loader --local-folder "C:\docs" --dry-run
python -m kb_loader --sharepoint-url "https://..." --existing update
python -m kb_loader --kb-status
python -m kb_loader --help
```

---

## Project structure

```
D365KBLoader/
├── run.bat                  ← Windows launcher (double-click)
├── run.command              ← Mac launcher (double-click)
├── requirements.txt         ← Python dependencies
├── README.md                ← this file
└── kb_loader/
    ├── __main__.py          ← CLI entry point
    ├── gui.py               ← Tkinter GUI
    ├── service.py           ← Core load logic (used by both CLI and GUI)
    ├── settings.py          ← User-profile settings store
    ├── auth.py              ← Microsoft authentication (MSAL + Azure CLI)
    ├── sharepoint_client.py ← SharePoint via Graph API
    ├── dataverse_client.py  ← D365 Knowledge Article CRUD
    ├── converter.py         ← Word → HTML conversion
    └── run_log.py           ← Excel run log generator
```
