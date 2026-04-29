# D365 Knowledge Base Loader

Cross-platform CLI tool that loads Word documents into Dynamics 365 Knowledge Base articles.

## What it does

1. Reads `.docx` files from a **local folder** or **SharePoint** (via Microsoft Graph)
2. Recursively processes all subfolders
3. Converts each document to HTML using [mammoth](https://github.com/mwilliamson/python-mammoth)
4. Saves HTML copies locally (preserving folder structure)
5. Creates Knowledge Articles in Dataverse with the HTML content
6. Publishes each article (Draft → Approved → Published)

## Prerequisites

- **Python 3.10+**
- **Azure CLI** (`az`) — [Install guide](https://aka.ms/installazurecli)
  - No app registration needed — just run `az login` with your Microsoft account
  - Your account needs access to the D365/Dataverse environment
  - For SharePoint mode: your account also needs access to the SharePoint site

## Setup

1. **Install Azure CLI** (if not already installed):
   - Windows: `winget install Microsoft.AzureCLI`
   - Mac: `brew install azure-cli`

2. **Install Python dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Configure:**
   ```bash
   # Windows
   copy .env.example .env

   # Mac/Linux
   cp .env.example .env
   ```
   Edit `.env` and set `DATAVERSE_URL=https://your-org.crm.dynamics.com`

## Usage

### Local folder (simplest — great for OneDrive-synced SharePoint folders)

```bash
python -m kb_loader --local-folder "C:\Users\you\OneDrive - Company\KB Articles"
```

### SharePoint folder (reads directly from SharePoint via Graph API)

```bash
python -m kb_loader --sharepoint-url "https://tenant.sharepoint.com/sites/MySite/Shared Documents/KB Articles"
```

### Dry run (convert to HTML only, don't upload to Dataverse)

```bash
python -m kb_loader --local-folder ./docs --dry-run
```

### All options

```bash
python -m kb_loader \
  --local-folder ./docs \
  --output-dir ./html_output \
  --existing skip \
  --verbose
```

## CLI Options

| Option | Description | Default |
|--------|-------------|---------|
| `--local-folder` | Local folder with .docx files | — |
| `--sharepoint-url` | SharePoint folder URL | — |
| `--output-dir` | Local directory for HTML files | `./output` |
| `--existing` | Handle duplicates: `skip`, `update`, or `duplicate` | `skip` |
| `--dry-run` | Convert only, don't upload to Dataverse | `false` |
| `--verbose` / `-v` | Enable debug logging | `false` |

## Authentication

The tool uses **Azure CLI** for authentication — no app registration required.

1. If you're not logged in, the tool automatically opens a browser for `az login`
2. It acquires separate tokens for Graph API (SharePoint) and Dataverse
3. Tokens are cached by Azure CLI across runs

## How it handles Knowledge Articles

Each article is created with:
- **Title**: Filename without extension (e.g., `Troubleshooting Guide.docx` → `Troubleshooting Guide`)
- **Content**: Full HTML converted from the Word document
- **Language**: English (locale 1033)
- **Creation mode**: Manual (`msdyn_creationmode = 0`)
- **Description**: Auto-generated with source path
- **Keywords**: Source file path (for traceability)
- **Status**: Published (transitions through Draft → Approved → Published)

## Idempotency

When re-running the tool, it checks for existing articles by title:
- **`skip`** (default): Skip files that already have a matching article
- **`update`**: Overwrite the existing article's content
- **`duplicate`**: Create a new article regardless

## Run Log

Each run generates a timestamped Excel file in the output directory (e.g. `output/kb_loader_log_20260429_143000.xlsx`) with:

| Column | Description |
|--------|-------------|
| File Name | Original .docx filename |
| Folder Path | Subfolder relative to input root |
| File Size (bytes) | Raw file size |
| Has Content | Whether HTML conversion produced content |
| HTML Saved | Whether the HTML file was saved locally |
| Published to KB | Whether it was published to Dataverse |
| KB Action | Created, Updated, Skipped, Dry Run, or Error |
| Article ID | Dataverse knowledge article ID |
| Error | Error message if processing failed |

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
    ├── converter.py          # Word → HTML conversion
    └── run_log.py            # Excel run log generator
```
