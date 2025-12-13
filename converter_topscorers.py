"""Amateurvoetbal topscorers: .txt/.docx -> Cue Web HTML-code (als tekstbestand).

Doel:
- Toon alle divisies/klassen als aparte secties.
- Maak per sectie een genummerde lijst (<ol>) met correcte <li>-items.
- Escape alle tekst correct (HTML-encoding) en gebruik <br> voor regels binnen één item.

Belangrijk voor Word (.docx):
- In het aangeleverde Word-document staat de rangnummering niet als '1.' in de tekst.
- In plaats daarvan begint een nieuw 'rank-blok' bij regels met '- <N> doelpunten'.
  (De daaropvolgende regels zonder doelpunten horen bij diezelfde rank.)

Deze converter volgt die logica en werkt ook met .txt-input die wel '1.' of '1 ' voor rangen bevat.
"""

from __future__ import annotations

import html as _html
import re
import tempfile
from typing import List, Tuple

from docx import Document


# =========================
# Templates (Cue Web)
# =========================

# Sectietitel (bijv. 'Eerste klasse', 'Derde en Vierde divisie')
HEADING_TEMPLATE = (
    '<h4 class="Heading_heading__okScq Heading_heading--sm__bGPWw heading_sm__u3F2n" '
    'data-testid="article-subhead">{title}</h4>'
)

# Ordered list: houd dit bewust "simpel" en stabiel; Cue Web kan dit in een HTML-blok renderen.
OL_OPEN = '<ol data-testid="numbered-list">'
OL_CLOSE = '</ol>'

LI_TEMPLATE = (
    '<li>'
    '<p class="Paragraph_paragraph__exhQA Paragraph_paragraph--default-sm-default__jy0uG articleParagraph">{content}</p>'
    '</li>'
)


# =========================
# Parsing helpers
# =========================

# Rangregel als die in txt al genummerd is: accepteer '1 ' en '1. '
NUMBER_RE = re.compile(r'^\s*(\d+)(?:\.)?\s+')

# Start van een nieuw rank-blok: "- 11 doelpunten"
GOALS_RE = re.compile(r'-\s*(\d+)\s*doelpunten\b', re.IGNORECASE)


def strip_source_rank_number(line: str) -> str:
    return re.sub(r'^\s*\d+(?:\.)?\s*', '', line, count=1)


def is_section_heading(text: str) -> bool:
    """Detecteer sectiekoppen.

    In jouw input zien we o.a.:
    - 'Derde en Vierde divisie'
    - 'Eerste klasse', 'Tweede klasse', etc.

    We willen niet per ongeluk spelersregels als heading markeren
    (die bevatten vaak '(' of '- N doelpunten').
    """
    t = (text or '').strip()
    if not t:
        return False

    if re.search(r'\b(klasse|klassen|divisie|divisies)\b', t, flags=re.IGNORECASE):
        if '(' in t:
            return False
        if GOALS_RE.search(t):
            return False
        if len(t) > 80:
            return False
        return True

    # Extra fallback: korte caps-headings (bv. 'DERDE DIVISIE B')
    letters = re.sub(r'[^A-Za-zÀ-ÿ]', '', t)
    if letters and len(t) <= 50:
        upper_ratio = sum(1 for ch in letters if ch.isupper()) / max(1, len(letters))
        if upper_ratio >= 0.8 and not '(' in t and not GOALS_RE.search(t):
            return True

    return False


def _escape_and_join_lines(lines: List[str]) -> str:
    safe = [_html.escape(ln.strip()) for ln in lines if ln.strip()]
    return '<br>'.join(safe)


# =========================
# DOCX / TXT extraction
# =========================

def extract_text_from_upload(raw: bytes, filename: str) -> str:
    """Lees upload (.txt of .docx) en geef de tekstinhoud terug."""
    name = (filename or '').lower()

    if name.endswith('.docx'):
        with tempfile.NamedTemporaryFile(suffix='.docx', delete=True) as tmp:
            tmp.write(raw)
            tmp.flush()
            doc = Document(tmp.name)

            # Neem alle paragrafen mee als losse regels.
            lines: List[str] = []
            for p in doc.paragraphs:
                txt = (p.text or '').strip()
                if txt:
                    lines.append(txt)

            # Tabellen (indien aanwezig)
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for ln in (cell.text or '').splitlines():
                            t = ln.strip()
                            if t:
                                lines.append(t)

            return '\n'.join(lines).strip()

    # default: txt
    return raw.decode('utf-8', errors='replace')


# =========================
# Main conversion
# =========================

def topscorers_text_to_cueweb_html(content: str) -> str:
    """Converteer topscorers-tekst naar Cue Web HTML-code (als tekst)."""
    raw_lines = [ln.rstrip() for ln in (content or '').splitlines()]

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
        if current_title:
            # ook als er (per ongeluk) geen items zijn, willen we de heading niet kwijt,
            # maar in praktijk verwachten we items.
            if current_items:
                sections.append((current_title, current_items))
            else:
                sections.append((current_title, []))
        current_title = None
        current_items = []

    for raw in raw_lines:
        line = (raw or '').strip()
        if not line:
            continue

        # Skip algemene titel bovenaan, als die vóór de eerste echte sectie staat
        if current_title is None and line.lower().startswith('tussenstand'):
            continue

        if is_section_heading(line):
            flush_group()
            flush_section()
            current_title = line
            continue

        # Start nieuw rank-blok:
        # 1) Als de bron expliciet genummerd is (txt)
        # 2) Of als er een '- N doelpunten' in de regel staat (docx/txt)
        if NUMBER_RE.match(line) or GOALS_RE.search(line):
            flush_group()
            current_group = [strip_source_rank_number(line)]
            continue

        # Subregel (zelfde rank)
        if not current_group:
            # fallback: begin een group als er nog geen rank gestart is
            current_group = [line]
        else:
            current_group.append(line)

    flush_group()
    flush_section()

    # Render HTML
    out: List[str] = []

    for title, items in sections:
        out.append(HEADING_TEMPLATE.format(title=_html.escape(title.strip())))

        if items:
            out.append(OL_OPEN)
            for group in items:
                out.append(LI_TEMPLATE.format(content=_escape_and_join_lines(group)))
            out.append(OL_CLOSE)
        else:
            # Geen items gevonden (zou zelden moeten); laat in elk geval de heading zien.
            out.append('<p class="Paragraph_paragraph__exhQA Paragraph_paragraph--default-sm-default__jy0uG articleParagraph"></p>')

        out.append('<br>')

    return '\n'.join(out).strip()
