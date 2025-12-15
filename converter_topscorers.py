#!/usr/bin/env python3
"""
converter_topscorers.py

Doel
-----
Zet een Word-bestand (.docx) of tekstbestand (.txt) met een topscorers-overzicht om
naar HTML. De output bestaat per klasse/divisie uit:

1) Een heading (HEADING_TEMPLATE)
2) Een genummerde lijst (TEMPLATE_HTML) met <li>-items waarin spelersregels staan.
   Gelijke standen worden binnen één <li> samengevoegd met "<br>\n".

Belangrijk
----------
- TEMPLATE_HTML en HEADING_TEMPLATE zijn 1-op-1 overgenomen uit het oorspronkelijke notebook.
- De HTML-strings worden NIET afgekort of vereenvoudigd.
- Sectiekoppen worden herkend op "klasse/divisie", maar spelerregels met "(..., ... divisie)"
  mogen niet als kop worden gezien (anders verdwijnt o.a. "Derde en Vierde divisie").

Dependency
----------
    pip install python-docx
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from typing import List, Tuple, Optional

import html as _html

try:
    from docx import Document  # type: ignore
except Exception:  # pragma: no cover
    Document = None  # type: ignore


# Ingebouwd HTML-sjabloon voor de genummerde lijst (EXACT uit notebook)
TEMPLATE_HTML = """<ol data-testid="numbered-list" class="List_list__TqiC5 List_list--ordered__jhPJG styles_list__7BMph styles_orderedList__wTCQI">

<li class="List_list-item__G_gHo">

<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" class="Icon_icon__SKejO Icon_icon--md__JEKjB List_list-item__icon__uA9Ih" aria-hidden="true"><path d="M12 15a3 3 0 1 0 0-6 3 3 0 0 0 0 6"></path></svg>

<p class="Paragraph_paragraph__exhQA Paragraph_paragraph--default-sm-default__jy0uG articleParagraph">speler 1</p>

</li>

<li class="List_list-item__G_gHo"><svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" class="Icon_icon__SKejO Icon_icon--md__JEKjB List_list-item__icon__uA9Ih" aria-hidden="true"><path d="M12 15a3 3 0 1 0 0-6 3 3 0 0 0 0 6"></path></svg><p class="Paragraph_paragraph__exhQA Paragraph_paragraph--default-sm-default__jy0uG articleParagraph">speler 2<br>
speler 3</p>

</li>

</ol>"""

# HTML-heading voor klasse/divisie-koppen (EXACT uit notebook)
HEADING_TEMPLATE = (
    '<h4 class="Heading_heading__okScq Heading_heading--sm__bGPWw '
    'heading_articleSubheading__HfjIx heading_sm__u3F2n" '
    'data-testid="article-subhead">{title}</h4>'
)

NUMBER_RE = re.compile(r"^\s*\d+\.\s")


def parse_html_template(template_text: str) -> Tuple[str, str, str]:
    """
    Splitst TEMPLATE_HTML in:
    - prefix: alles t/m opening <ol> + eventuele whitespace vóór de eerste <li> inhoud
    - item_template: een volledig <li>..</li> met {content} op de plek van <p>inhoud
    - suffix: rest t/m </ol>
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
    p_open, _, p_close = m_p.groups()

    item_template = li_block[: m_p.start()] + p_open + "{content}" + p_close + li_block[m_p.end():]
    return prefix, item_template, suffix


def looks_like_player_stat_line(line: str) -> bool:
    """
    Notebook-logica:
    - spelerregels bevatten vaak (TEAM, ... divisie) en/of '- X doelpunten'
    Deze moeten NIET als sectiekop worden gezien.
    """
    s = line.strip()
    lower = s.lower()
    if "(" in s and ")" in s:
        return True
    if "-" in s and re.search(r"\b\d+\b", s) and "doelpunt" in lower:
        return True
    return False


def is_section_heading(line: str) -> bool:
    """
    Notebook-logica (aangepast aan docx output):
    - niet leeg
    - niet genummerd (1., 2., ...)
    - bevat 'klasse' of 'divisie'
    - maar lijkt niet op spelerregel
    """
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


def extract_text_lines_from_docx(path: Path) -> List[str]:
    if Document is None:
        raise RuntimeError("python-docx is niet beschikbaar. Installeer met: pip install python-docx")

    doc = Document(str(path))
    lines: List[str] = []

    # paragrafen
    for p in doc.paragraphs:
        lines.append(p.text)

    # tabellen (voor de zekerheid)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                lines.extend(cell.text.splitlines())

    # behoud lege regels als scheiding, maar strip \r\n
    return [l.rstrip("\n") for l in lines]


def extract_text_lines_from_txt(path: Path) -> List[str]:
    raw = path.read_bytes()
    try:
        text = raw.decode("utf-8")
    except UnicodeDecodeError:
        text = raw.decode("cp1252")
    return text.splitlines()


def parse_sections_from_lines(lines: List[str]) -> List[Tuple[str, List[List[str]]]]:
    """
    Notebook-structuur:
    - sectiekoppen (klasse/divisie)
    - binnen sectie: groepen gescheiden door lege regels
      waarbij een groep 1 of meer regels bevat (gelijke stand -> meerdere regels)
    """
    sections: List[Tuple[str, List[List[str]]]] = []
    current_title: Optional[str] = None
    current_groups: List[List[str]] = []
    current_group: List[str] = []
    started = False

    def flush_group():
        nonlocal current_group
        if current_group:
            # strip eventueel "1. " etc, maar behoud inhoud verder intact
            cleaned: List[str] = []
            for l in current_group:
                if NUMBER_RE.match(l):
                    cleaned.append(strip_source_rank_number(l))
                else:
                    cleaned.append(l)
            current_groups.append(cleaned)
            current_group = []

    def flush_section():
        nonlocal current_groups, current_title
        if current_title and current_groups:
            sections.append((current_title, current_groups))
        current_groups = []

    for raw in lines:
        line = (raw or "").rstrip("\n").strip()

        if is_section_heading(line):
            flush_group()
            flush_section()
            current_title = line
            started = True
            continue

        if not started:
            # alles vóór de eerste sectiekop negeren (documenttitel e.d.)
            continue

        if not line:
            flush_group()
            continue

        current_group.append(line)

    flush_group()
    flush_section()
    return sections


def apply_template(template_text: str, sections: List[Tuple[str, List[List[str]]]]) -> str:
    prefix, item_template, suffix = parse_html_template(template_text)
    html_parts: List[str] = []

    for title, groups in sections:
        html_parts.append(HEADING_TEMPLATE.format(title=_html.escape(title)))

        items: List[str] = []
        for group in groups:
            # Escape per regel; gelijkstand: <br>\n
            safe_lines = [_html.escape(l, quote=False) for l in group]
            inner = "<br>\n".join(safe_lines)
            items.append(item_template.replace("{content}", inner))

        html_parts.append(prefix + "\n" + "\n\n".join(items) + "\n" + suffix)

    return "\n\n".join(html_parts)


def convert_file(input_path: Path) -> str:
    if not input_path.exists():
        raise FileNotFoundError(f"Bestand bestaat niet: {input_path}")

    suffix = input_path.suffix.lower()
    if suffix == ".docx":
        lines = extract_text_lines_from_docx(input_path)
    else:
        lines = extract_text_lines_from_txt(input_path)

    sections = parse_sections_from_lines(lines)
    if not sections:
        raise ValueError(
            "Geen secties gevonden. Verwacht minstens één regel met 'klasse' of 'divisie' als kop."
        )

    return apply_template(TEMPLATE_HTML, sections)


def build_arg_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Converteer topscorers-stand (.docx/.txt) naar HTML in een tekstbestand."
    )
    p.add_argument("input", help="Pad naar inputbestand (.docx of .txt)")
    p.add_argument(
        "-o", "--output",
        help="Pad naar outputbestand (default: <input_stem>_output_html.txt)"
    )
    return p


def main(argv: Optional[List[str]] = None) -> int:
    args = build_arg_parser().parse_args(argv)

    in_path = Path(args.input).expanduser().resolve()
    html_out = convert_file(in_path)

    out_path = Path(args.output).expanduser().resolve() if args.output else in_path.with_name(
        f"{in_path.stem}_output_html.txt"
    )

    out_path.write_text(html_out, encoding="utf-8")
    print(f"Gereed: {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
