"""Amateurvoetbal topscorers: tekstbestand (.txt/.docx) -> Cue Web HTML-code.

Gebaseerd op notebook:
`12 klassement_html_converter_v12 - integratie docx.ipynb`.

De output is HTML-code die als tekst (.txt) wordt aangeboden zodat deze eenvoudig te kopiëren/plakken is.
"""

from __future__ import annotations

import html as _html
import re
import tempfile
from typing import Tuple, List

from docx import Document

# HTML-heading voor klasse/divisie-koppen (Cue Web)
HEADING_TEMPLATE = '<h4 class="Heading_heading__okScq Heading_heading--sm__bGPWw heading_articleSubheading__HfjIx heading_sm__u3F2n" data-testid="article-subhead">{title}</h4>'

# Ingebouwd HTML-sjabloon voor de genummerde lijst.
# We gebruiken een stabiele (minimale) variant zonder hashed classnamen voor de <ol>/<li>,
# maar behouden de Paragraph-classes voor de inhoudsregels.
TEMPLATE_HTML = '<ol data-testid="numbered-list">\n<li>\n<p class="Paragraph_paragraph__exhQA Paragraph_paragraph--default-sm-default__jy0uG articleParagraph">{content}</p>\n</li>\n</ol>'

NUMBER_RE = re.compile(r"^\s*\d+\.\s")


def parse_html_template(template_text: str) -> Tuple[str, str, str]:
    """Splits een <ol> template op in: prefix, item_template, suffix.

    item_template bevat {content} als placeholder voor de inhoud.
    """
    m_ol = re.search(r"(<ol[^>]*>)(.*?)(</ol>)", template_text, re.S)
    if not m_ol:
        raise ValueError("Kon geen <ol>...</ol> in het sjabloon vinden.")
    prefix = template_text[: m_ol.start(2)]
    suffix = template_text[m_ol.end(2):]

    m_li = re.search(r"(<li\b.*?</li>)", m_ol.group(2), re.S)
    if not m_li:
        raise ValueError("Kon geen <li> in het sjabloon vinden.")
    li_block = m_li.group(1)

    m_p = re.search(r"(<p\b[^>]*>)(.*?)(</p>)", li_block, re.S)
    if not m_p:
        raise ValueError("Kon geen <p> in het <li>-sjabloon vinden.")

    item_template = li_block[: m_p.start(2)] + "{content}" + li_block[m_p.end(2):]
    return prefix, item_template, suffix


def looks_like_player_stat_line(line: str) -> bool:
    s = line.strip()
    lower = s.lower()
    if "(" in s and ")" in s:
        return True
    if "-" in s and re.search(r"\b\d+\b", s) and "doelpunt" in lower:
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
    if looks_like_player_stat_line(s):
        return False
    return True


def strip_source_rank_number(line: str) -> str:
    return re.sub(r"^\s*\d+\.\s*", "", line, count=1)


def parse_sections(text: str):
    """Parseer inputtekst naar secties (titel + groepen regels).

    Herkent headings met 'KLASSE' of 'DIVISIE'. Een regel '1. ...' start een nieuwe speler-groep.
    """
    lines = text.splitlines()
    sections = []
    current_title = None
    current_groups: List[List[str]] = []
    current_group: List[str] = []

    def flush_group():
        nonlocal current_group, current_groups
        if current_group:
            current_groups.append(current_group)
            current_group = []

    def flush_section():
        nonlocal current_title, current_groups, sections
        if current_title and current_groups:
            sections.append((current_title, current_groups))
        current_groups = []

    for raw_line in lines:
        line = raw_line.rstrip("\n").strip()
        if not line:
            continue

        if is_section_heading(line):
            flush_group()
            flush_section()
            current_title = line.strip()
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


def topscorers_text_to_cueweb_html(text: str) -> str:
    """Converteer topscorers tekst naar Cue Web HTML-code."""
    prefix, item_template, suffix = parse_html_template(TEMPLATE_HTML)
    sections = parse_sections(text)
    html_parts = []

    for title, groups in sections:
        html_parts.append(HEADING_TEMPLATE.format(title=_html.escape(title)))
        items = []
        for group in groups:
            safe_lines = [_html.escape(l, quote=False) for l in group]
            inner = "<br>\n".join(safe_lines)
            items.append(item_template.replace("{content}", inner))
        html_parts.append(prefix + "\n" + "\n\n".join(items) + "\n" + suffix)

    return "\n\n".join(html_parts).strip() + "\n"


def extract_text_from_upload(raw: bytes, filename: str) -> str:
    """Lees upload (.txt of .docx) en geef de tekstinhoud terug."""
    name = (filename or "").lower()
    if name.endswith(".docx"):
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=True) as tmp:
            tmp.write(raw)
            tmp.flush()
            doc = Document(tmp.name)
            lines = []
            for p in doc.paragraphs:
                if p.text.strip():
                    lines.append(p.text)
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for ln in cell.text.splitlines():
                            if ln.strip():
                                lines.append(ln)
            return "\n".join(lines)

    # default: txt
    return raw.decode("utf-8", errors="replace")


# ------------------------------------------------------------
# Compat: bytes-uploader API (gebruikt door app.py)
# ------------------------------------------------------------
def extract_text_from_upload_bytes(raw: bytes, filename: str) -> str:
    """Lees upload-bytes en geef tekst terug.
    Wrapper voor extract_text_from_upload(...), zodat app.py een stabiele API heeft.
    """
    return extract_text_from_upload(raw, filename)

