"""Convert Word documents (.docx) to HTML using mammoth."""

import io
import logging
from pathlib import Path
import mammoth

logger = logging.getLogger(__name__)


def convert_docx_to_html(docx_bytes: bytes) -> tuple[str, list[str]]:
    """Convert a .docx file (as bytes) to HTML.

    Returns:
        (html_content, warnings) - the HTML string and any conversion warnings.
    """
    with io.BytesIO(docx_bytes) as f:
        result = mammoth.convert_to_html(f)

    if result.messages:
        for msg in result.messages:
            logger.warning(f"Conversion warning: {msg}")

    warnings = [str(m) for m in result.messages]
    return result.value, warnings


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
