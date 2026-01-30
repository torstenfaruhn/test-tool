#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
converter_topscorers.py (Render-compatible build)

Key fix for your Render error:
- Exposes `convert_topscorers_upload(file_bytes, filename)` so an app can import:
  `from converter_topscorers import convert_topscorers_upload`

Also includes a CLI entrypoint.

Dependencies:
- python-docx (only needed for .docx input)

HTML templates are copied in full and are not abbreviated.
"""

from __future__ import annotations

from pathlib import Path
import argparse
import html as _html
import re
from typing import List, Tuple, Optional

try:
    from docx import Document  # python-docx
except Exception:
    Document = None  # type: ignore


# ---------------------------------------------------------------------------
# HTML templates (MUST remain complete and unabridged)
# ---------------------------------------------------------------------------

TEMPLATE_HTML = """<ol data-testid="numbered-list" class="List_list__TqiC5 List_list--ordered__GX7eF styles_list__7BMph styles_orderedList__wTCQI">

<li class="List_list-item__KWec8">

<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" class="Icon_icon__IIDV_ Icon_icon--md__hvhhZ List_list-item__icon__z7Lcl" aria-hidden="true"><path d="M12 15a3 3 0 1 0 0-6 3 3 0 0 0 0 6"></path></svg>

<p class="Paragraph_paragraph__fI0_V Paragraph_paragraph--default-sm-default__1te_f articleParagraph">speler 1</p>

</li>

<li class="List_list-item__KWec8"><svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" class="Icon_icon__IIDV_ Icon_icon--md__hvhhZ List_list-item__icon__z7Lcl" aria-hidden="true"><path d="M12 15a3 3 0 1 0 0-6 3 3 0 0 0 0 6"></path></svg><p class="Paragraph_paragraph__fI0_V Paragraph_paragraph--default-sm-default__1te_f articleParagraph">speler 2<br>
speler 3</p>

</li>

</ol>"""

HEADING_TEMPLATE = (
    '<h4 class="Heading_heading__u5NOQ Heading_heading--sm__yIwgk '
    'heading_articleSubheading__HfjIx heading_sm__u3F2n" '
    'data-testid="article-subhead">{title}</h4>'
)


# ---------------------------------------------------------------------------
# Parsing logic (ported from the notebook; prevents "divisie" inside player lines
# from being mistaken as section headings)
# ---------------------------------------------------------------------------

NUMBER_RE = re.compile(r"^\s*\d+\.\s")
GOALS_RE = re.compile(r"\b\d+\b")
DOELPUNT_RE = re.compile(r"\bdoelpunt", re.IGNORECASE)


def parse_html_template(template_text: str) -> Tuple[str, str, str]:
    m_ol = re.search(r"(<ol[^>]*>)(.*?)(</ol>)", template_text, re.S)
    if not m_ol:
        raise ValueError("Kon geen <ol>...</ol> in het sjabloon vinden.")
    prefix = template_text[: m_ol.start(2)]
    suffix = template_text[m_ol.end(2) :]

    m_li = re.search(r"(<li\b.*?</li>)", m_ol.group(2), re.S)
    if not m_li:
        raise ValueError("Kon geen <li> in het sjabloon vinden.")
    li_block = m_li.group(1)

    m_p = re.search(r"(<p\b[^>]*>)(.*?)(</p>)", li_block, re.S)
    if not m_p:
        raise ValueError("Kon geen <p> in het <li>-sjabloon vinden.")
    p_open, _, p_close = m_p.groups()

    item_template = (
        li_block[: m_p.start()] + p_open + "{content}" + p_close + li_block[m_p.end() :]
    )
    return prefix, item_template, suffix


def looks_like_player_stat_line(line: str) -> bool:
    s = line.strip()
    lower = s.lower()

    if "(" in s and ")" in s:
        return True

    if "-" in s and GOALS_RE.search(s) and DOELPUNT_RE.search(lower):
        return True

    return False


def is_section_heading(line: str) -> bool:
    s = line.strip()
    if not s:
        return False
    if NUMBER_RE.match(s):
        return False

    upper = s.upper()
    if "KLASSE" not in upper and "DIVISIE" not in upper:
        return False

    # Crucial: player lines like "... (.., vierde divisie) - 14 doelpunten"
    # must not be treated as headings.
    if looks_like_player_stat_line(s):
        return False

    return True


def strip_source_rank_number(line: str) -> str:
    return re.sub(r"^\s*\d+\.\s*", "", line, count=1)


def parse_sections(text: str) -> List[Tuple[str, List[List[str]]]]:
    lines = text.splitlines()
    sections: List[Tuple[str, List[List[str]]]] = []

    current_title: Optional[str] = None
    current_groups: List[List[str]] = []
    current_group: List[str] = []

    def flush_group() -> None:
        nonlocal current_group
        if current_group:
            current_groups.append(current_group)
            current_group = []

    def flush_section() -> None:
        nonlocal current_groups, current_title
        if current_title and current_groups:
            sections.append((current_title, current_groups))
        current_groups = []

    for raw_line in lines:
        line = raw_line.rstrip("\n")

        if is_section_heading(line):
            flush_group()
            flush_section()
            current_title = line.strip()
            continue

        if not line.strip():
            flush_group()
            continue

        if NUMBER_RE.match(line):
            flush_group()
            current_group = [strip_source_rank_number(line)]
        else:
            if not current_group:
                current_group = [line]
            else:
                current_group.append(line)

    flush_group()
    flush_section()
    return sections


def apply_template(template_text: str, klassement_text: str) -> str:
    prefix, item_template, suffix = parse_html_template(template_text)
    sections = parse_sections(klassement_text)
    html_parts: List[str] = []

    for title, groups in sections:
        html_parts.append(HEADING_TEMPLATE.format(title=_html.escape(title)))
        items: List[str] = []
        for group in groups:
            safe_lines = [_html.escape(l, quote=False) for l in group]
            inner = "<br>\n".join(safe_lines)
            items.append(item_template.replace("{content}", inner))
        html_parts.append(prefix + "\n" + "\n\n".join(items) + "\n" + suffix)

    return "\n\n".join(html_parts)


# ---------------------------------------------------------------------------
# Input extraction
# ---------------------------------------------------------------------------

def extract_text_from_docx_bytes(file_bytes: bytes) -> str:
    if Document is None:
        raise ImportError(
            "python-docx is niet beschikbaar. Voeg 'python-docx' toe aan je dependencies."
        )

    tmp_path = Path("_uploaded_input.docx")
    tmp_path.write_bytes(file_bytes)

    doc = Document(str(tmp_path))
    lines: List[str] = []

    for p in doc.paragraphs:
        lines.append(p.text)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                lines.extend(cell.text.splitlines())

    return "\n".join(lines)


def extract_text_from_upload_bytes(file_bytes: bytes, filename: str) -> str:
    name = (filename or "").lower()

    if name.endswith(".docx"):
        return extract_text_from_docx_bytes(file_bytes)

    try:
        return file_bytes.decode("utf-8")
    except UnicodeDecodeError:
        return file_bytes.decode("cp1252")


# ---------------------------------------------------------------------------
# Public API for web apps (Render import)
# ---------------------------------------------------------------------------

def convert_topscorers_upload(file_bytes: bytes, filename: str) -> str:
    """Convert an uploaded file (bytes + original filename) to HTML."""
    text = extract_text_from_upload_bytes(file_bytes, filename)
    return apply_template(TEMPLATE_HTML, text)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def convert_file_to_html(in_path: Path) -> str:
    return convert_topscorers_upload(in_path.read_bytes(), in_path.name)


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(
        description="Converteer topscoorders/klassement (.docx of .txt) naar HTML."
    )
    parser.add_argument("input", help="Pad naar inputbestand (.docx of .txt).")
    parser.add_argument(
        "-o",
        "--output",
        default="klassement_output_html.txt",
        help="Pad naar outputbestand (tekstbestand met HTML).",
    )

    args = parser.parse_args(argv)
    in_path = Path(args.input)
    out_path = Path(args.output)

    out_html = convert_file_to_html(in_path)
    out_path.write_text(out_html, encoding="utf-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
