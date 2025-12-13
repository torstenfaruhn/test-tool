"""Amateurvoetbal topscorers: tekstbestand (.txt/.docx) -> Cue Web HTML-code (als .txt).

Deze versie hardcodeert de Cue Web-markup voor:
- tussenkop (klasse/divisie)
- numbered list (<ol>/<li>)

De templates zijn overgenomen uit het notebook:
`12 klassement_html_converter_v12 - integratie docx.ipynb`
(en zijn daarom bewust niet “versimpeld” naar generieke HTML).

Input: Word of tekst met sectiekoppen (klasse/divisie) en regels met "- N doelpunten".
Output: per sectie een heading + een genummerde lijst met items.
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

# NB: In het notebook staan in de ol/svg-class delen met "..." (afkorting).
# Cue Web herkent de lijst op basis van data-testid + de List_* classes.
# Daarom hardcoderen we hier de volledige (niet-afgekorte) class voor de <ol>.
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
# Parsing heuristieken
# -----------------------------

# Start van een item: bevat "- N doelpunt(en)" (zoals in je Word/kladblok bron)
POINTS_RE = re.compile(r"\s-\s*\d+\s+doelpunt(?:en)?\s*$", flags=re.IGNORECASE)

def _is_section_heading(line: str) -> bool:
    t = (line or "").strip()
    if not t:
        return False

    # Veelvoorkomend: "Eerste klasse", "Tweede klasse", etc.
    if re.search(r"\bklasse\b", t, flags=re.IGNORECASE):
        return True

    # Divisies (vaak caps): "DERDE DIVISIE B", "TWEEDE DIVISIE", etc.
    if re.search(r"\bdivisie\b", t, flags=re.IGNORECASE):
        return True

    # Extra heuristiek voor korte CAPS-koppen zonder cijfers
    if not re.search(r"\d", t):
        letters = re.sub(r"[^A-Za-zÀ-ÿ]", "", t)
        if letters and len(t) <= 45:
            upper_ratio = sum(1 for ch in letters if ch.isupper()) / max(1, len(letters))
            if upper_ratio >= 0.8:
                return True

    return False


def extract_text_from_upload(raw: bytes, filename: str) -> List[str]:
    """Lees upload (.txt of .docx) en geef regels terug.

    DOCX: neem paragrafen + tabellen. We vertrouwen niet op Word-lijstnummering;
    de item-start wordt afgeleid uit '- N doelpunten' in de tekst.
    """
    name = (filename or "").lower()

    if name.endswith(".docx"):
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=True) as tmp:
            tmp.write(raw)
            tmp.flush()
            doc = Document(tmp.name)

            lines: List[str] = []
            for p in doc.paragraphs:
                t = (p.text or "").strip()
                if t:
                    lines.append(t)

            # Tabellen (indien gebruikt)
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
    return [ln.strip() for ln in text.splitlines() if ln.strip()]


def topscorers_text_to_cueweb_html(lines: List[str]) -> str:
    """Maak Cue Web HTML-code (als tekst) uit lijst met regels."""
    sections: List[Tuple[str, List[List[str]]]] = []

    current_title: str | None = None
    current_items: List[List[str]] = []
    current_item: List[str] = []

    def flush_item():
        nonlocal current_item, current_items
        if current_item:
            current_items.append(current_item)
            current_item = []

    def flush_section():
        nonlocal current_title, current_items
        if current_title and current_items:
            sections.append((current_title, current_items))
        current_title = None
        current_items = []

    for line in lines:
        if _is_section_heading(line):
            flush_item()
            flush_section()
            current_title = line.strip()
            continue

        # Nieuwe rank start: regel eindigt op "- N doelpunten"
        if POINTS_RE.search(line):
            flush_item()
            current_item = [line.strip()]
        else:
            # subregel (meerdere namen bij dezelfde rank)
            if not current_item:
                # Fallback: als er (nog) geen item is, start er toch één
                current_item = [line.strip()]
            else:
                current_item.append(line.strip())

    flush_item()
    flush_section()

    # Render
    out: List[str] = []
    for title, items in sections:
        out.append(HEADING_TEMPLATE.format(title=_html.escape(title)))

        out.append(OL_OPEN)
        for item_lines in items:
            # escape + <br>
            safe_lines = [_html.escape(x) for x in item_lines if x.strip()]
            content = "<br>\n".join(safe_lines)
            out.append(LI_TEMPLATE.format(content=content))
        out.append(OL_CLOSE)
        out.append("<br>")  # spacer tussen secties

    return "\n".join(out).strip()


# Convenience wrapper voor app.py
def convert_topscorers_upload(raw: bytes, filename: str) -> str:
    lines = extract_text_from_upload(raw, filename)
    return topscorers_text_to_cueweb_html(lines)
