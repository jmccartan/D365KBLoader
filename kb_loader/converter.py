"""Convert Word documents (.docx and .doc) to HTML using mammoth.

For legacy .doc files, LibreOffice is used to convert to .docx first.
"""

import io
import logging
import subprocess
import shutil
import tempfile
from pathlib import Path
import mammoth

logger = logging.getLogger(__name__)

SUPPORTED_EXTENSIONS = {".docx", ".doc"}


def _find_libreoffice() -> str | None:
    """Find the LibreOffice executable on this system."""
    for name in ("libreoffice", "soffice"):
        path = shutil.which(name)
        if path:
            return path
    # Common install locations
    common_paths = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ]
    for p in common_paths:
        if Path(p).is_file():
            return p
    return None


def _convert_doc_to_docx(doc_bytes: bytes) -> bytes:
    """Convert a legacy .doc file to .docx using LibreOffice headless."""
    lo_path = _find_libreoffice()
    if not lo_path:
        raise RuntimeError(
            "LibreOffice is required to convert .doc files but was not found.\n"
            "Install it from: https://www.libreoffice.org/download/\n"
            "After installing, ensure 'soffice' is in your PATH."
        )

    with tempfile.TemporaryDirectory() as tmpdir:
        doc_path = Path(tmpdir) / "input.doc"
        doc_path.write_bytes(doc_bytes)

        result = subprocess.run(
            [lo_path, "--headless", "--convert-to", "docx", "--outdir", tmpdir, str(doc_path)],
            capture_output=True, text=True, timeout=120,
        )

        if result.returncode != 0:
            raise RuntimeError(f"LibreOffice conversion failed: {result.stderr.strip()}")

        docx_path = Path(tmpdir) / "input.docx"
        if not docx_path.exists():
            raise RuntimeError("LibreOffice conversion produced no output file.")

        return docx_path.read_bytes()


def convert_to_html(file_bytes: bytes, file_name: str) -> tuple[str, list[str]]:
    """Convert a .docx or .doc file (as bytes) to HTML with inline D365 styles.

    For .doc files, converts to .docx via LibreOffice first, then uses mammoth.
    The mammoth output is then post-processed to add inline styles compatible
    with the D365 Knowledge Article rich-text editor (which strips <style>
    blocks).

    Returns:
        (html_content, warnings) - the styled HTML string and any conversion warnings.
    """
    ext = Path(file_name).suffix.lower()

    if ext == ".doc":
        logger.info("  Converting .doc → .docx via LibreOffice...")
        file_bytes = _convert_doc_to_docx(file_bytes)

    with io.BytesIO(file_bytes) as f:
        result = mammoth.convert_to_html(f)

    if result.messages:
        for msg in result.messages:
            logger.warning(f"Conversion warning: {msg}")

    raw_html = result.value
    warnings = [str(m) for m in result.messages]

    # Apply consistent inline styling for D365 KB articles
    try:
        from kb_loader.styles import style_html
        styled = style_html(raw_html)
    except Exception as e:
        logger.warning(f"Could not apply inline styles: {e} — using raw mammoth output")
        styled = raw_html

    return styled, warnings


def save_html_file(html: str, output_dir: str, relative_path: str, original_name: str) -> Path:
    """Save HTML content to a file, preserving the SharePoint folder structure.

    Args:
        html: The HTML content to save.
        output_dir: Root output directory.
        relative_path: Subfolder path relative to the SharePoint root folder.
        original_name: Original .docx filename.

    Returns:
        Path to the saved HTML file.
    """
    html_name = Path(original_name).stem + ".html"
    out_path = Path(output_dir) / relative_path / html_name
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(html, encoding="utf-8")
    logger.info(f"Saved HTML: {out_path}")
    return out_path
