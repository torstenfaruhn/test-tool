"""Amateurvoetbal online: Cue Print (txt) -> Word (DOCX)

Deze module is gebaseerd op het notebook:
`03 colab_tag_converter-docx.ipynb`.

Output-regels:
- <subhead_lead> (divisie/klasse) komt 1x als kop (UPPERCASE + bold)
- Daarna volgen alle wedstrijden (<subhead>) als eigen alinea.
  - Facts (<howto_facts>) komen in dezelfde alinea op de volgende regel (Shift+Enter) en zijn italic.
  - Lege facts-blokken: overslaan.
- Tussen competitieblokken: 1 lege alinea.
- Geen logging van inhoud; alleen technische counters.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

from docx import Document
from docx.enum.text import WD_BREAK

# Matcht: <subhead_lead>...</subhead_lead>, <subhead>...</subhead>, <howto_facts>...</howto_facts>
_TAG_PATTERN = re.compile(
    r"<(subhead_lead|subhead|howto_facts)>(.*?)</\1>",
    re.DOTALL | re.IGNORECASE,
)


@dataclass(frozen=True)
class Token:
    kind: str  # subhead_lead, subhead, howto_facts
    text: str  # inhoud binnen de tag


def _extract_tokens(raw: str) -> List[Token]:
    tokens: List[Token] = []
    for m in _TAG_PATTERN.finditer(raw):
        kind = (m.group(1) or "").lower().strip()
        text = (m.group(2) or "").replace("\r\n", "\n").replace("\r", "\n").strip()
        tokens.append(Token(kind=kind, text=text))
    return tokens


@dataclass(frozen=True)
class Item:
    header: str
    subhead: str
    facts: Optional[str]  # None als leeg/ontbrekend


def _tokens_to_items(tokens: List[Token]) -> Tuple[List[Item], int]:
    """Zet tokens om naar items (wedstrijdregels), met header-context (divisie/klasse).

    Returns:
      - items: lijst met wedstrijd-items
      - block_changes: aantal keer dat een nieuwe header start (t.o.v. vorige)
    """
    items: List[Item] = []
    current_header: Optional[str] = None
    block_changes = 0
    i = 0

    while i < len(tokens):
        t = tokens[i]

        if t.kind == "subhead_lead":
            if current_header is not None and t.text != current_header:
                block_changes += 1
            current_header = t.text
            i += 1
            continue

        if t.kind == "subhead":
            # Zonder header-context geen output (input is dan onverwacht)
            if not current_header:
                i += 1
                continue

            subhead_text = t.text
            facts_text: Optional[str] = None

            # Facts staan (meestal) direct achter subhead
            if i + 1 < len(tokens) and tokens[i + 1].kind == "howto_facts":
                candidate = (tokens[i + 1].text or "").strip()
                if candidate:
                    facts_text = candidate
                i += 2
            else:
                i += 1

            items.append(Item(header=current_header, subhead=subhead_text, facts=facts_text))
            continue

        i += 1

    return items, block_changes


def cueprint_txt_to_docx(content: str, output_path: str) -> Dict[str, int]:
    """Zet Cue Print-code (txt als string) om naar een Word-document (.docx).

    output_path: pad waar de .docx wordt opgeslagen.

    Returns: technische stats (geen inhoud).
    """
    tokens = _extract_tokens(content)
    items, _ = _tokens_to_items(tokens)

    doc = Document()

    stats: Dict[str, int] = {
        "tokens_total": len(tokens),
        "headers_total": 0,
        "items_total": 0,
        "items_with_facts": 0,
        "empty_facts_skipped": 0,
        "block_separators_added": 0,
    }

    prev_header: Optional[str] = None

    for it in items:
        header = (it.header or "").strip()

        # Nieuw competitieblok?
        if header and header != prev_header:
            if prev_header is not None:
                doc.add_paragraph("")  # 1 lege alinea tussen competitieblokken
                stats["block_separators_added"] += 1

            hp = doc.add_paragraph()
            hr = hp.add_run(header.upper())
            hr.bold = True
            stats["headers_total"] += 1
            prev_header = header

        # Wedstrijdregel als eigen alinea
        p = doc.add_paragraph()
        r1 = p.add_run((it.subhead or "").strip())

        facts = (it.facts or "").strip()
        if facts:
            r1.add_break(WD_BREAK.LINE)  # Shift+Enter
            r2 = p.add_run(facts)
            r2.italic = True
            stats["items_with_facts"] += 1
        else:
            stats["empty_facts_skipped"] += 1

        stats["items_total"] += 1

    doc.save(output_path)
    return stats
