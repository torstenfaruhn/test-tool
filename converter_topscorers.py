"""Amateurvoetbal topscorers: tekstbestand (.txt/.docx) -> Cue Web HTML-code.

Gebaseerd op notebook:
`12 klassement_html_converter_v12 - integratie docx.ipynb`.

De output is HTML-code die als tekst (.txt) wordt aangeboden zodat deze eenvoudig te kopiëren/plakken is.

Deze versie herstelt twee issues die in de repository-variant problemen gaven:
- Geen afgekorte '...' class strings in templates.
- Robuuste nummerherkenning (met/zonder punt) én betere DOCX-extractie (Word-nummering).
"""

from __future__ import annotations

import html as _html
import io
import re
import tempfile
from typing import List, Tuple

from docx import Document


# =========================
# Templates (Cue Web)
# =========================

# Kop boven een sectie (bijv. "Eerste klasse")
# NB: In Cue Web worden deze classes vaak als "hashed" gebruikt; we nemen hier een volledige (niet-afgekorte) variant op.
HEADING_TEMPLATE = (
    '<h4 class="Heading_heading__okScq Heading_heading--sm__bGPWw heading_sm__u3F2n" '
    'data-testid="article-subhead">{title}</h4>'
)

# Genummerde lijst template. We houden <ol>/<li> minimaal (stabiel), en gebruiken de volledige Paragraph classes.
TEMPLATE_HTML = (
    '<ol data-testid="numbered-list">\n'
    '<li>\n'
    '<p class="Paragraph_paragraph__exhQA Paragraph_paragraph--default-sm-default__jy0uG articleParagraph">{content}</p>\n'
    '</li>\n'
    '</ol>'
)


def parse_html_template(template_text: str) -> Tuple[str, str, str]:
    """Splits een <ol> template op in: prefix, item_template, suffix.

    item_template bevat {content} als placeholder voor de inhoud.
    """
    m_ol = re.search(r"(<ol[^>]*>)(.*?)(</ol>)", template_text, re.S | re.I)
    if not m_ol:
        raise ValueError("Kon geen <ol>...</ol> in het sjabloon vinden.")

    prefix = template_text[: m_ol.start(2)]
    suffix = template_text[m_ol.end(2) :]

    m_li = re.search(r"(<li\b.*?</li>)", m_ol.group(2), re.S | re.I)
    if not m_li:
        raise ValueError("Kon geen <li>...</li> in het sjabloon vinden.")

    item_template = m_li.group(1)
    if "{content}" not in item_template:
        raise ValueError("Het <li>-sjabloon bevat geen {content} placeholder.")

    return prefix, item_template, suffix


_OL_PREFIX, _ITEM_TEMPLATE, _OL_SUFFIX = parse_html_template(TEMPLATE_HTML)


# =========================
# Parsing helpers
# =========================

# Rangregel: accepteer '1 ' en '1. ' (met/zonder punt)
NUMBER_RE = re.compile(r"^\s*\d+(?:\.)?\s+")


def strip_source_rank_number(line: str) -> str:
    """Verwijder een eventueel rangnummer aan het begin van de regel (met/zonder punt)."""
    return re.sub(r"^\s*\d+(?:\.)?\s*", "", line, count=1)


def is_section_heading(text: str) -> bool:
    """Heuristiek: detecteer sectiekoppen zoals 'Eerste klasse', 'Tweede klasse', etc."""
    t = (text or "").strip()
    if not t:
        return False

    # Veelvoorkomend: 'Eerste klasse', 'Tweede klasse', 'Derde klasse', etc.
    if re.search(r"\bklasse\b", t, flags=re.I):
        return True

    # Ook koppen als 'DERDE DIVISIE B' (caps) moeten als heading worden gezien.
    # Heuristiek: geen leidend nummer én relatief kort én veel uppercase.
    if not NUMBER_RE.match(t):
        letters = re.sub(r"[^A-Za-zÀ-ÿ]", "", t)
        if letters and len(t) <= 40:
            upper_ratio = sum(1 for ch in letters if ch.isupper()) / max(1, len(letters))
            if upper_ratio >= 0.8:
                return True

    return False


def _escape_and_join_lines(lines: List[str]) -> str:
    """Escape tekst en zet multi-line content om naar HTML met <br>."""
    safe = [_html.escape(ln.strip()) for ln in lines if ln.strip()]
    return "<br>".join(safe)


# =========================
# DOCX / TXT extraction
# =========================

def extract_text_from_upload(raw: bytes, filename: str) -> str:
    """Lees upload (.txt of .docx) en geef de tekstinhoud terug.

    Belangrijk: bij DOCX staat het lijstnummer vaak niet in paragraph.text.
    We detecteren daarom Word-numbering en voegen een '1. ' prefix toe per sectie.
    """
    name = (filename or "").lower()

    if name.endswith(".docx"):
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=True) as tmp:
            tmp.write(raw)
            tmp.flush()
            doc = Document(tmp.name)

            lines: List[str] = []
            list_counter = 0

            for p in doc.paragraphs:
                text = (p.text or "").strip()
                if not text:
                    continue

                # Sectiekop -> reset numbering
                if is_section_heading(text):
                    list_counter = 0
                    lines.append(text)
                    continue

                # Detecteer Word numbering: numPr aanwezig op pPr
                has_numbering = (
                    getattr(p, "_p", None) is not None
                    and getattr(p._p, "pPr", None) is not None
                    and getattr(p._p.pPr, "numPr", None) is not None
                )

                if has_numbering:
                    list_counter += 1
                    lines.append(f"{list_counter}. {text}")
                else:
                    lines.append(text)

            # Tabellen (indien gebruikt)
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for ln in (cell.text or "").splitlines():
                            t = ln.strip()
                            if t:
                                lines.append(t)

            return "\n".join(lines).strip()

    # default: txt
    return raw.decode("utf-8", errors="replace")


# =========================
# Main conversion
# =========================

def topscorers_text_to_cueweb_html(content: str) -> str:
    """Converteer topscorers-tekst naar Cue Web HTML-code (als tekst).

    Verwacht input met sectiekoppen en genummerde regels.
    Subregels (zonder nummer) worden aan de vorige rang toegevoegd.
    """
    lines = [ln.rstrip() for ln in (content or "").splitlines()]
    sections: List[Tuple[str, List[List[str]]]] = []

    current_title: str | None = None
    current_items: List[List[str]] = []
    current_group: List[str] = []

    def flush_group() -> None:
        nonlocal current_group, current_items
        if current_group:
            current_items.append(current_group)
            current_group = []

    def flush_section() -> None:
        nonlocal current_title, current_items
        if current_title and current_items:
            sections.append((current_title, current_items))
        current_title = None
        current_items = []

    for raw_line in lines:
        line = (raw_line or "").strip()
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
                # Als er geen actieve groep is, begin een groep (fallback)
                current_group = [line]
            else:
                current_group.append(line)

    flush_group()
    flush_section()

    # Render HTML
    out: List[str] = []
    for title, items in sections:
        out.append(HEADING_TEMPLATE.format(title=_html.escape(title.strip())))

        # ordered list
        out.append(_OL_PREFIX)
        for group in items:
            content_html = _escape_and_join_lines(group)
            out.append(_ITEM_TEMPLATE.format(content=content_html))
        out.append(_OL_SUFFIX)

        # kleine spacer tussen secties
        out.append("<br>")

    return "\n".join(out).strip()
