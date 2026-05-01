# =====================================================================
# D365 Knowledge Base Loader
# Copyright (c) 2026 John McCartan
# Licensed under the MIT License. See the LICENSE file in the project
# root for the full text.
# =====================================================================

"""Inline-style post-processor for HTML produced by mammoth.

D365 Knowledge Article rich-text editor strips `<style>` blocks and external
stylesheets, so all formatting must live in `style="..."` attributes on each
element. This module walks the HTML produced by mammoth and applies a
consistent Microsoft-Fluent-inspired theme inline.

Also detects common prose patterns and lifts them into styled callouts:
  - paragraphs starting with "Note:", "Warning:", or "Tip:" become callout boxes
"""

from __future__ import annotations

from bs4 import BeautifulSoup, NavigableString


# ── Style definitions (all inline-friendly) ────────────────────────────────

WRAPPER_STYLE = (
    "font-family:'Segoe UI',Arial,sans-serif; font-size:14px; "
    "line-height:1.5; color:#201f1e; max-width:820px;"
)

H1_STYLE = (
    "font-size:22px; color:#0078d4; border-bottom:1px solid #edebe9; "
    "padding-bottom:6px; margin:0 0 12px 0;"
)
H2_STYLE = "font-size:18px; color:#0078d4; margin:20px 0 8px 0;"
H3_STYLE = "font-size:15px; color:#323130; margin:16px 0 6px 0;"
H4_STYLE = "font-size:14px; color:#323130; margin:14px 0 4px 0; font-weight:600;"
H5_STYLE = "font-size:13px; color:#605e5c; margin:12px 0 4px 0; font-weight:600;"
H6_STYLE = (
    "font-size:12px; color:#605e5c; margin:10px 0 4px 0; "
    "font-weight:600; text-transform:uppercase;"
)

P_STYLE = "margin:8px 0;"
A_STYLE = "color:#0078d4; text-decoration:none;"

UL_STYLE = "margin:8px 0 8px 24px; padding:0;"
OL_STYLE = "margin:8px 0 8px 24px; padding:0;"
LI_STYLE = "margin:4px 0;"

CODE_STYLE = (
    "font-family:Consolas,monospace; background:#f3f2f1; padding:1px 4px; "
    "border-radius:3px; font-size:13px;"
)
PRE_STYLE = (
    "background:#f3f2f1; padding:10px; border-left:3px solid #0078d4; "
    "font-family:Consolas,monospace; font-size:13px; white-space:pre-wrap;"
)

TABLE_STYLE = "border-collapse:collapse; width:100%; margin:10px 0;"
TH_STYLE = (
    "border:1px solid #edebe9; padding:6px 10px; text-align:left; "
    "background:#faf9f8; font-weight:600;"
)
TD_STYLE = "border:1px solid #edebe9; padding:6px 10px; vertical-align:top;"

# Callout styles
CALLOUT_NOTE_STYLE = (
    "background:#fff4ce; border-left:3px solid #ffb900; "
    "padding:8px 12px; margin:10px 0;"
)
CALLOUT_WARNING_STYLE = (
    "background:#fde7e9; border-left:3px solid #d13438; "
    "padding:8px 12px; margin:10px 0;"
)
CALLOUT_TIP_STYLE = (
    "background:#dff6dd; border-left:3px solid #107c10; "
    "padding:8px 12px; margin:10px 0;"
)


# Map element name → style; applied if the element has no inline style yet
_TAG_STYLES = {
    "h1": H1_STYLE,
    "h2": H2_STYLE,
    "h3": H3_STYLE,
    "h4": H4_STYLE,
    "h5": H5_STYLE,
    "h6": H6_STYLE,
    "p": P_STYLE,
    "a": A_STYLE,
    "ul": UL_STYLE,
    "ol": OL_STYLE,
    "li": LI_STYLE,
    "code": CODE_STYLE,
    "pre": PRE_STYLE,
    "table": TABLE_STYLE,
    "th": TH_STYLE,
    "td": TD_STYLE,
}

# Callout prefixes (case-insensitive)
_CALLOUT_PATTERNS = (
    ("note:", CALLOUT_NOTE_STYLE, "Note"),
    ("warning:", CALLOUT_WARNING_STYLE, "Warning"),
    ("important:", CALLOUT_WARNING_STYLE, "Important"),
    ("caution:", CALLOUT_WARNING_STYLE, "Caution"),
    ("tip:", CALLOUT_TIP_STYLE, "Tip"),
    ("hint:", CALLOUT_TIP_STYLE, "Hint"),
)


def _merge_style(existing: str, addition: str) -> str:
    """Append `addition` to `existing` style attribute. Existing properties win."""
    if not existing:
        return addition
    if not addition:
        return existing
    # Naive merge: existing first, then addition. CSS rule order means
    # addition wins for duplicate properties — but in our use case existing
    # will only contain mammoth-set things like `text-align:right` from Word
    # paragraph alignment, which we want to preserve.
    sep = "" if existing.rstrip().endswith(";") else "; "
    return f"{existing}{sep}{addition}"


def _apply_tag_styles(soup: BeautifulSoup):
    """Walk the soup and add inline styles to known tags."""
    for tag_name, style in _TAG_STYLES.items():
        for el in soup.find_all(tag_name):
            existing = el.get("style", "")
            el["style"] = _merge_style(existing, style)


def _convert_callouts(soup: BeautifulSoup):
    """Detect 'Note:', 'Warning:', 'Tip:' paragraphs and convert to callout divs."""
    for p in list(soup.find_all("p")):
        # Get the leading text content (skip whitespace nodes)
        first = None
        for child in p.children:
            if isinstance(child, NavigableString) and not str(child).strip():
                continue
            first = child
            break
        if first is None:
            continue

        # Get the leading text
        if isinstance(first, NavigableString):
            leading_text = str(first).lstrip()
        elif first.name in ("strong", "b", "em", "i"):
            leading_text = first.get_text().lstrip()
        else:
            continue

        lower = leading_text.lower()
        match = next(((p_, s_, l_) for p_, s_, l_ in _CALLOUT_PATTERNS if lower.startswith(p_)), None)
        if not match:
            continue

        prefix, style, label = match

        # Build new callout div
        new_div = soup.new_tag("div", style=style)
        strong = soup.new_tag("strong")
        strong.string = f"{label}:"
        new_div.append(strong)

        # Strip the prefix from the leading text and append remaining content
        if isinstance(first, NavigableString):
            remaining = str(first)[len(leading_text) - len(leading_text.lstrip()):].lstrip()
            # Strip the actual prefix (case-preserving)
            # The actual prefix length is len(prefix) (without the colon? prefix already has colon)
            # Drop the matched prefix from the original string
            stripped_idx = lower.find(prefix) + len(prefix)
            new_text = leading_text[stripped_idx:]
            if new_text:
                new_div.append(NavigableString(" " + new_text.lstrip()))
            first.extract()
        else:
            # leading element is a strong/em tag containing "Note:" etc — drop it
            first.extract()
            new_div.append(NavigableString(" "))

        # Move all remaining children of the original <p> into the callout div
        for child in list(p.children):
            new_div.append(child.extract())

        p.replace_with(new_div)


def _wrap_in_themed_div(soup: BeautifulSoup) -> str:
    """Wrap the body content in a themed outer div and return the HTML string."""
    # mammoth output has no <html>/<body>, just bare elements
    inner = soup.encode(formatter="html").decode("utf-8")
    return f'<div style="{WRAPPER_STYLE}">\n{inner}\n</div>'


def style_html(html: str) -> str:
    """Apply consistent inline styles to mammoth-produced HTML.

    Returns an HTML string suitable for direct paste into D365 Knowledge
    Article rich-text content (no <style> blocks, all styling inline).
    """
    if not html or not html.strip():
        return html

    soup = BeautifulSoup(html, "html.parser")

    # Convert prose patterns first (they replace <p> with <div>)
    _convert_callouts(soup)

    # Then apply tag styles to everything that's left
    _apply_tag_styles(soup)

    return _wrap_in_themed_div(soup)
