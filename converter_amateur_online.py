"""Amateurvoetbal online: Cue Print (txt) -> Cue Web (HTML-code)

Deze module is gebaseerd op het notebook:
`02 colab_tag_converter_with_br_spacing.ipynb`.

Optie 1: volledige classnamen (zoals in het werkende notebook / gewenste output).
"""

from __future__ import annotations

import re


_REPLACEMENTS: dict[str, str] = {
    "<body>": "",
    "</body>": "",
    "<subhead_lead>": (
        '<h5 class="Heading_heading__tL6MO Heading_heading--lg__wP2Ux heading_lg__W3ya6" '
        'data-testid="infoblock-headline">'
    ),
    "</subhead_lead>": "</h5>",
    "<subhead>": (
        '<h6 class="Heading_heading__tL6MO Heading_heading--sm__n8pqT heading_infoblockSubheading__Ecn_I heading_sm__u3F2n" '
        'data-testid="infoblock-heading">'
    ),
    "</subhead>": "</h6>",
    "<howto_facts>": (
        '<p class="Paragraph_paragraph__exhQA Paragraph_paragraph--default-sm-default__jy0uG articleParagraph">'
    ),
    "</howto_facts>": "</p><br>",
}


def cueprint_txt_to_cueweb_html(content: str) -> str:
    """Zet Cue Print-code (txt) om naar Cue Web HTML-code (als tekst)."""
    for old, new in _REPLACEMENTS.items():
        content = content.replace(old, new)

    # Extra witruimte: wanneer een facts-paragraaf direct gevolgd wordt door een headline,
    # voeg één extra <br> toe.
    content = re.sub(
        r"(</p><br>\s*)(<h5\b[^>]*data-testid=\"infoblock-headline\"[^>]*>)",
        r"\1<br>\2",
        content,
        flags=re.IGNORECASE,
    )

    return content
