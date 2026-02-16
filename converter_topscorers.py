"""Amateurvoetbal topscorers: tekstbestand (.txt/.docx) -> Word (.docx) voor Cue Web.

Gebaseerd op notebook:
`12 klassement_html_converter_v12 - integratie docx.ipynb`.

Output: een Word-document met per divisie/klasse een kop (bold) en een genummerde lijst
waarbij de nummering per sectie opnieuw start bij 1. Alleen het lijstnummer ("1.") is bold.
Spelers met gelijk aantal doelpunten staan onder hetzelfde nummer met een soft line break
(Shift+Enter).
"""

from __future__ import annotations

import re
import tempfile
from io import BytesIO
from typing import List, Tuple

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


NUMBER_RE = re.compile(r"^\s*\d+\.\s*")


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
    """Verwijder een bron-rangnummer ('1. ') uit een regel."""
    return NUMBER_RE.sub("", line).strip()


def parse_sections(text: str) -> List[Tuple[str, List[List[str]]]]:
    """Parseer inputtekst naar secties (titel + groepen regels).

    - Heading: regels met 'KLASSE' of 'DIVISIE'
    - Een regel '1. ...' start een nieuwe speler-groep.
    - Regels erna horen bij dezelfde groep (gelijk aantal goals) en krijgen een soft line break.
    """
    lines = text.splitlines()
    sections: List[Tuple[str, List[List[str]]]] = []
    current_title: str | None = None
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


# -----------------------------
# DOCX numbering (robust)
# -----------------------------

def _add_abstract_numbering_number_bold(document: Document) -> int:
    """Maak een abstract numbering definitie met decimal numbering en bold nummering.

    Retourneert abstractNumId.
    """
    numbering = document.part.numbering_part.numbering_definitions._numbering  # type: ignore[attr-defined]

    existing_ids = []
    for an in numbering.findall(qn("w:abstractNum")):
        try:
            existing_ids.append(int(an.get(qn("w:abstractNumId"))))
        except Exception:
            pass
    abstract_id = (max(existing_ids) + 1) if existing_ids else 1

    abstract = OxmlElement("w:abstractNum")
    abstract.set(qn("w:abstractNumId"), str(abstract_id))

    # single level (0)
    lvl = OxmlElement("w:lvl")
    lvl.set(qn("w:ilvl"), "0")

    start = OxmlElement("w:start")
    start.set(qn("w:val"), "1")
    lvl.append(start)

    numFmt = OxmlElement("w:numFmt")
    numFmt.set(qn("w:val"), "decimal")
    lvl.append(numFmt)

    lvlText = OxmlElement("w:lvlText")
    lvlText.set(qn("w:val"), "%1.")
    lvl.append(lvlText)

    # Nummering bold (alleen de "1.")
    rPr = OxmlElement("w:rPr")
    b = OxmlElement("w:b")
    b.set(qn("w:val"), "1")
    rPr.append(b)
    lvl.append(rPr)

    # minimale inspringing, vergelijkbaar met Word default list
    pPr = OxmlElement("w:pPr")
    ind = OxmlElement("w:ind")
    ind.set(qn("w:left"), "720")   # 0.5 inch
    ind.set(qn("w:hanging"), "360")  # 0.25 inch
    pPr.append(ind)
    lvl.append(pPr)

    abstract.append(lvl)
    numbering.append(abstract)
    return abstract_id


def _add_numbering_instance(document: Document, abstract_num_id: int) -> int:
    """Maak een nieuwe numbering instance (numId) voor een sectie, start altijd bij 1."""
    numbering = document.part.numbering_part.numbering_definitions._numbering  # type: ignore[attr-defined]

    existing_ids = []
    for n in numbering.findall(qn("w:num")):
        try:
            existing_ids.append(int(n.get(qn("w:numId"))))
        except Exception:
            pass
    num_id = (max(existing_ids) + 1) if existing_ids else 1

    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(num_id))

    abstract_ref = OxmlElement("w:abstractNumId")
    abstract_ref.set(qn("w:val"), str(abstract_num_id))
    num.append(abstract_ref)

    numbering.append(num)
    return num_id


def _apply_num_id(paragraph, num_id: int, level: int = 0) -> None:
    """Koppel een paragraph aan een numbering instance."""
    p = paragraph._p  # lxml element
    pPr = p.get_or_add_pPr()
    numPr = pPr.find(qn("w:numPr"))
    if numPr is None:
        numPr = OxmlElement("w:numPr")
        pPr.append(numPr)

    ilvl = numPr.find(qn("w:ilvl"))
    if ilvl is None:
        ilvl = OxmlElement("w:ilvl")
        numPr.append(ilvl)
    ilvl.set(qn("w:val"), str(level))

    numId = numPr.find(qn("w:numId"))
    if numId is None:
        numId = OxmlElement("w:numId")
        numPr.append(numId)
    numId.set(qn("w:val"), str(num_id))


def topscorers_text_to_docx_bytes(text: str) -> bytes:
    """Converteer topscorers tekst naar een Word-document (.docx) en geef bytes terug."""
    sections = parse_sections(text)

    doc = Document()

    abstract_id = _add_abstract_numbering_number_bold(doc)

    for title, groups in sections:
        # Sectiekop: bold, style Normal
        p_head = doc.add_paragraph()
        run = p_head.add_run(title)
        run.bold = True

        # Nummering per sectie opnieuw starten
        num_id = _add_numbering_instance(doc, abstract_id)

        for group in groups:
            # één genummerde alinea per groep
            p_item = doc.add_paragraph(style="Body Text")
            _apply_num_id(p_item, num_id, level=0)

            if not group:
                continue

            # eerste regel
            p_item.add_run(group[0])

            # extra regels: soft line break (Shift+Enter)
            for extra in group[1:]:
                r = p_item.add_run()
                r.add_break()  # line break binnen dezelfde alinea
                p_item.add_run(extra)

            # spacing zoals voorbeeld: lege regel na elk item
            doc.add_paragraph("", style="Body Text")

        # extra witregel na de sectie
        doc.add_paragraph("")

    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()


# -----------------------------
# Upload text extraction
# -----------------------------

def extract_text_from_upload(raw: bytes, filename: str) -> str:
    """Lees upload (.txt of .docx) en geef de tekstinhoud terug."""
    name = (filename or "").lower()
    if name.endswith(".docx"):
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=True) as tmp:
            tmp.write(raw)
            tmp.flush()
            doc = Document(tmp.name)
            lines: List[str] = []
            for p in doc.paragraphs:
                if p.text.strip():
                    lines.append(p.text)
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            if p.text.strip():
                                lines.append(p.text)
            return "\n".join(lines).strip() + "\n"
    else:
        try:
            text = raw.decode("utf-8")
        except UnicodeDecodeError:
            text = raw.decode("latin-1")
        return text.strip() + "\n"


def extract_text_from_upload_bytes(raw: bytes, filename: str) -> str:
    """Lees upload-bytes en geef tekst terug.
    Wrapper voor extract_text_from_upload(...), zodat app.py een stabiele API heeft.
    """
    return extract_text_from_upload(raw, filename)
