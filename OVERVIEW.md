# D365 Knowledge Base Loader — Overview

## What Is It?

A command-line tool that bulk-loads Word documents into Dynamics 365 Knowledge Base articles. It converts `.docx` files to HTML and publishes them as knowledge articles in your D365 environment.

Works on both **Windows** and **Mac**.

## How It Works

1. **Point it at a folder** of Word documents (local folder or SharePoint URL)
2. It recursively finds all `.docx` files, including subfolders
3. Each document is converted to clean HTML
4. An HTML copy is saved locally for reference
5. A Knowledge Article is created in D365 Dataverse with:
   - **Title** = the filename (without `.docx`)
   - **Content** = the converted HTML
   - **Status** = Published
   - **Creation mode** = Manual
   - **Language** = English
6. The article is automatically transitioned through Draft → Approved → Published

## Authentication

Uses **Azure CLI** — no app registration or admin setup required.

- Install Azure CLI (`az`), then run `az login`
- The tool handles the rest, including opening a browser sign-in if needed
- Your Microsoft account just needs access to the D365 environment (and SharePoint, if using that mode)

## Two Input Modes

| Mode | Command | Best For |
|------|---------|----------|
| **Local folder** | `--local-folder "C:\path\to\docs"` | OneDrive-synced SharePoint folders, or any local `.docx` files |
| **SharePoint URL** | `--sharepoint-url "https://..."` | Reading directly from SharePoint without syncing locally |

## Quick Start

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Set your Dataverse URL
cp .env.example .env
# Edit .env → set DATAVERSE_URL=https://your-org.crm.dynamics.com

# 3. Run it
python -m kb_loader --local-folder "C:\Users\you\Documents\KB Articles"
```

## Key Features

- **Dry run mode** (`--dry-run`) — convert files to HTML without uploading, useful for previewing
- **Idempotent** — on re-runs, choose to `skip`, `update`, or `duplicate` existing articles
- **Resilient** — retries on API throttling (429) and server errors (5xx)
- **Preserves structure** — HTML output mirrors the source folder hierarchy
- **Traceability** — stores the source file path in the article's keywords and description

## Requirements

- Python 3.10+
- Azure CLI (`az`)
- Access to a D365 Customer Service environment with Knowledge Management enabled
