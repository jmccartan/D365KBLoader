# D365 Knowledge Base Loader

A cross-platform command-line tool that bulk-loads Word documents (`.docx` and `.doc`) into Dynamics 365 Knowledge Base articles. Works on both **Windows** and **Mac**.

## How It Works

1. **Point it at a folder** of Word documents — either a local folder or a SharePoint URL
2. It recursively finds all `.docx` and `.doc` files, including subfolders
3. Each document is converted to clean HTML (using [mammoth](https://github.com/mwilliamson/python-mammoth); legacy `.doc` files are first converted via [LibreOffice](https://www.libreoffice.org/))
4. An HTML copy is saved locally for reference, preserving the folder structure
5. A Knowledge Article is created in D365 Dataverse with:
   - **Title** = the filename (without extension)
   - **Content** = the converted HTML
   - **Status** = Published (transitions through Draft → Approved → Published)
   - **Creation mode** = Manual
   - **Language** = English
6. An Excel log file is generated for each run summarizing every file processed

## Key Features

- **Two input modes** — read from a local folder or directly from SharePoint
- **No app registration required** — authenticates via Azure CLI (`az login`)
- **Dry run mode** (`--dry-run`) — convert files to HTML without uploading, useful for previewing
- **Idempotent** — on re-runs, choose to `skip`, `update`, or `duplicate` existing articles
- **Resilient** — retries on API throttling (429) and server errors (5xx)
- **Traceability** — stores the source file path in each article's keywords and description
- **Run log** — timestamped Excel report for every run

---

## Prerequisites

- **Python 3.10+**
  - Windows: [Download from python.org](https://www.python.org/downloads/) or `winget install Python.Python.3.13`
  - Mac: `brew install python`
- **Azure CLI** (`az`) — [Install guide](https://aka.ms/installazurecli)
  - Windows: `winget install Microsoft.AzureCLI`
  - Mac: `brew install azure-cli`
  - No app registration or admin setup needed — just sign in with your Microsoft account
  - Your account needs access to the D365/Dataverse environment
  - For SharePoint mode: your account also needs access to the SharePoint site
- **LibreOffice** (only needed if you have legacy `.doc` files — not required for `.docx`)
  - The tool will notify you with install instructions if a `.doc` file is encountered and LibreOffice is not installed
  - Windows: [Download from libreoffice.org](https://www.libreoffice.org/download/) or `winget install TheDocumentFoundation.LibreOffice`
  - Mac: `brew install --cask libreoffice`
- **D365 Customer Service** environment with Knowledge Management enabled

---

## Setup

### 1. Install Python dependencies

```bash
pip install -r requirements.txt
```

### 2. Configure your Dataverse URL

**Windows:**
```cmd
copy .env.example .env
```

**Mac:**
```bash
cp .env.example .env
```

Open `.env` in a text editor and set your Dataverse environment URL:
```
DATAVERSE_URL=https://your-org.crm.dynamics.com
```

### 3. Log in to Azure (one-time)

```bash
az login
```

This opens a browser where you sign in with your Microsoft account. Tokens are cached across runs.

---

## Usage

> **Recommended:** Run with `--dry-run` first to convert files and review the HTML output before publishing to Dataverse.

### Step 1: Dry run (preview — converts to HTML only, nothing is published)

**Windows:**
```cmd
python -m kb_loader --local-folder "C:\Users\you\OneDrive - Company\KB Articles" --dry-run
```

**Mac:**
```bash
python -m kb_loader --local-folder ~/OneDrive\ -\ Company/KB\ Articles --dry-run
```

Review the HTML files in the `./output` folder and check the Excel run log to confirm the right files were picked up.

### Step 2: Publish to Dataverse (when you're happy with the dry run)

**Windows:**
```cmd
python -m kb_loader --local-folder "C:\Users\you\OneDrive - Company\KB Articles"
```

**Mac:**
```bash
python -m kb_loader --local-folder ~/OneDrive\ -\ Company/KB\ Articles
```

### SharePoint folder (reads directly from SharePoint via Graph API)

```bash
python -m kb_loader --sharepoint-url "https://tenant.sharepoint.com/sites/MySite/Shared Documents/KB Articles"
```

### All options combined

**Windows:**
```cmd
python -m kb_loader --local-folder "C:\docs" --output-dir "C:\html_output" --existing skip --verbose
```

**Mac:**
```bash
python -m kb_loader --local-folder ./docs --output-dir ./html_output --existing skip --verbose
```

---

## CLI Options

| Option | Description | Default |
|--------|-------------|---------|
| `--local-folder` | Local folder with Word files (.docx, .doc) | — |
| `--sharepoint-url` | SharePoint folder URL | — |
| `--output-dir` | Local directory for HTML files and run logs | `./output` |
| `--existing` | Handle duplicates: `skip`, `update`, or `duplicate` | `skip` |
| `--dry-run` | Convert only, don't upload to Dataverse | `false` |
| `--verbose` / `-v` | Enable debug logging | `false` |

> **Note:** `--local-folder` and `--sharepoint-url` are mutually exclusive — use one or the other.

---

## Authentication

The tool uses **Azure CLI** for authentication — no app registration required.

1. If you're not logged in, the tool automatically opens a browser for `az login`
2. It acquires separate tokens for Graph API (SharePoint access) and Dataverse (article creation)
3. Tokens are cached by Azure CLI across runs — you won't need to log in every time

---

## Knowledge Article Details

Each article is created in Dataverse with these fields:

| Field | Value |
|-------|-------|
| **Title** | Filename without extension (e.g., `Troubleshooting Guide.docx` → `Troubleshooting Guide`) |
| **Content** | Full HTML converted from the Word document |
| **Language** | English (locale 1033) |
| **Creation mode** | Manual (`msdyn_creationmode = 0`) |
| **Description** | Auto-generated with source file path |
| **Keywords** | Source file path (for traceability) |
| **Status** | Published (transitions through Draft → Approved → Published) |

---

## Idempotency

When re-running the tool, it checks for existing articles by title:

| Mode | Behavior |
|------|----------|
| **`skip`** (default) | Skip files that already have a matching article |
| **`update`** | Overwrite the existing article's content |
| **`duplicate`** | Create a new article regardless |

---

## Run Log

Each run generates a timestamped Excel file in the output directory (e.g. `output/kb_loader_log_20260429_143000.xlsx`).

The log includes a summary header and a row per file with these columns:

| Column | Description |
|--------|-------------|
| File Name | Original filename |
| Folder Path | Subfolder relative to input root |
| File Size (bytes) | Raw file size |
| Has Content | Whether HTML conversion produced content |
| HTML Saved | Whether the HTML file was saved locally |
| Published to KB | Whether it was published to Dataverse |
| KB Action | Created, Updated, Skipped, Dry Run, or Error |
| Article ID | Dataverse knowledge article ID |
| Error | Error message if processing failed |

Cells are color-coded (green = yes, red = no, yellow = skipped) and the sheet includes auto-filter and frozen headers.

---

## Project Structure

```
D365KBLoader/
├── requirements.txt          # Python dependencies
├── .env.example              # Configuration template
├── README.md                 # This file
└── kb_loader/
    ├── __init__.py
    ├── __main__.py           # CLI entry point
    ├── config.py             # Configuration from .env
    ├── auth.py               # Azure CLI authentication
    ├── sharepoint_client.py  # SharePoint file enumeration & download
    ├── dataverse_client.py   # Knowledge Article CRUD & publishing
    ├── converter.py          # Word → HTML conversion (.docx + .doc)
    └── run_log.py            # Excel run log generator
```
