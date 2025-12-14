"""Amateurvoetbal topscorers: .docx/.txt -> Cue Web HTML-code (geleverd als .txt).

Doel
- Converteer een Word/Kladblok topscorers-overzicht naar Cue Web HTML-code.
- Cue Web herkent de “numbered list” alleen met specifieke classes/data-testid.
  Daarom zijn de tussenkop- en list-templates hieronder hardcoded overgenomen uit
  `12 klassement_html_converter_v12 - integratie docx.ipynb` (geen versimpeling).

Belangrijkste fixes t.o.v. eerdere repo-versies
- Alle secties (divisies + alle klassen) worden meegenomen: geen ‘stil’ wegvallen na 2 secties.
- Lijst-items worden herkend op basis van het patroon “- N doelpunt(en)”, niet op “1.” nummering
  (Word-nummering zit vaak niet in paragraph.text).
- Numbered list markup blijft intact (ol/li/p classes + data-testid).
- Alle inhoud wordt HTML-escaped; regels binnen één item worden met <br> samengevoegd.
"""

from __future__ import annotations

import html as _html
import re
import tempfile
from typing import List, Tuple

from docx import Document


# -----------------------------
# Hardcoded Cue Web templates (uit notebook)
# -----------------------------
HEADING_TEMPLATE = (
    '<h4 class="Heading_heading__okScq Heading_heading--sm__bGPWw '
    'heading_articleSubheading__HfjIx heading_sm__u3F2n" '
    'data-testid="article-subhead">{title}</h4>'
)

OL_OPEN = (
    '<ol data-testid="numbered-list" '
    'class="List_list__TqiC5 List_list--ordered__jhPJG styles_list__7BMph styles_orderedList__wTCQI">'
)
OL_CLOSE = "</ol>"

LI_TEMPLATE = (
    '<li class="List_list-item__G_gHo">'
    '<p class="Paragraph_paragraph__exhQA Paragraph_paragraph--default-sm-default__jy0uG articleParagraph">{content}</p>'
    "</li>"
)

# -----------------------------
# Parsing regex / heuristiek
# -----------------------------

# Item-start: eindigt (meestal) op “- N doelpunt(en)”
# We maken dit bewust tolerant: doelpunt/doelpunten, met/zonder extra spaties.
POINTS_RE = re.compile(r"\s-\s*\d+\s+doelpunt(?:en)?\s*$", flags=re.IGNORECASE)

# Sectiekop: bevat “klasse” of “divisie” (ook combinaties zoals “Derde en Vierde divisie”)
SECTION_RE = re.compile(r"\b(klasse|divisie)\b", flags=re.IGNORECASE)


def is_section_heading(line: str) -> bool:
    t = (line or "").strip()
    if not t:
        return False

    if SECTION_RE.search(t):
        return True

    # Extra: korte CAPS koppen (soms gebruikt voor divisies zonder het woord 'divisie')
    # bv. "DERDE DIVISIE B" valt al onder SECTION_RE; deze is extra defensief.
    if not re.search(r"\d", t):
        letters = re.sub(r"[^A-Za-zÀ-ÿ]", "", t)
        if letters and len(t) <= 50:
            upper_ratio = sum(1 for ch in letters if ch.isupper()) / max(1, len(letters))
            if upper_ratio >= 0.85:
                return True

    return False


def extract_lines_from_upload(raw: bytes, filename: str) -> List[str]:
    """Lees upload (.txt of .docx) en retourneer een genormaliseerde lijst regels."""
    name = (filename or "").lower()

    lines: List[str] = []

    if name.endswith(".docx"):
        # docx is zip; daarom niet blokkeren op PK...
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=True) as tmp:
            tmp.write(raw)
            tmp.flush()
            doc = Document(tmp.name)

            for p in doc.paragraphs:
                t = (p.text or "").strip()
                if t:
                    lines.append(t)

            # Tabellen (voor het geval de bron in een tabel staat)
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for ln in (cell.text or "").splitlines():
                            tt = ln.strip()
                            if tt:
                                lines.append(tt)

        return lines

    # default: txt
    text = raw.decode("utf-8", errors="replace")
    for ln in text.splitlines():
        t = (ln or "").strip()
        if t:
            lines.append(t)
    return lines


def build_sections(lines: List[str]) -> List[Tuple[str, List[List[str]]]]:
    """Zet regels om naar secties: (titel, [ [itemregel1, subregel2, ...], ... ])"""
    sections: List[Tuple[str, List[List[str]]]] = []

    current_title: str | None = None
    current_items: List[List[str]] = []
    current_item: List[str] = []

    def flush_item() -> None:
        nonlocal current_item, current_items
        if current_item:
            current_items.append(current_item)
            current_item = []

    def flush_section() -> None:
        nonlocal current_title, current_items
        if current_title is not None:
            # Voeg ook lege secties toe? In praktijk beter: alleen als er items zijn.
            if current_items:
                sections.append((current_title, current_items))
        current_title = None
        current_items = []

    for line in lines:
        # Sla een algemene titel over, als die bovenaan staat (bijv. "Tussenstand Topscorers")
        # Dit voorkomt een "lege sectie" die later verwarring kan geven.
        if current_title is None and not current_items and not current_item:
            if not is_section_heading(line) and line.lower().startswith("tussenstand"):
                continue

        if is_section_heading(line):
            flush_item()
            flush_section()
            current_title = line.strip()
            continue

        # Binnen sectie: item-start obv "- N doelpunten"
        if POINTS_RE.search(line):
            flush_item()
            current_item = [line.strip()]
        else:
            if not current_item:
                # Fallback: als de bron onvoorzien is, start toch een item
                current_item = [line.strip()]
            else:
                current_item.append(line.strip())

    flush_item()
    flush_section()

    return sections


def render_sections_to_cueweb_html(sections: List[Tuple[str, List[List[str]]]]) -> str:
    out: List[str] = []

    for title, items in sections:
        out.append(HEADING_TEMPLATE.format(title=_html.escape(title)))

        out.append(OL_OPEN)
        for item_lines in items:
            # escape + <br>
            safe = [_html.escape(x) for x in item_lines if x.strip()]
            content = "<br>\n".join(safe)
            out.append(LI_TEMPLATE.format(content=content))
        out.append(OL_CLOSE)

        out.append("<br>")  # spacer tussen secties

    return "\n".join(out).strip()


def convert_topscorers_upload(raw: bytes, filename: str) -> str:
    """Wrapper voor app.py."""
    lines = extract_lines_from_upload(raw, filename)
    sections = build_sections(lines)
    return render_sections_to_cueweb_html(sections)
