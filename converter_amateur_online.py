"""Amateurvoetbal online: Cue Print (txt) -> Cue Web (HTML-code)

Deze module is gebaseerd op het notebook:
`02 colab_tag_converter_with_br_spacing.ipynb`.

De oorspronkelijke Colab-specifieke bestandsupload/download is hier vervangen door
een pure functie die een string omzet.
"""

from __future__ import annotations

import re


# Replacements overgenomen uit het notebook (met ... als 'wildcard-achtige' class-aanduiding).
_REPLACEMENTS: dict[str, str] = {
    "<body>": "",
    "</body>": "",
    "<subhead_lead>": (
        '<h5 class="Heading_heading__okScq Heading...lg__v_Lob heading_lg__W3ya6" '
        'data-testid="infoblock-headline">'
    ),
    "</subhead_lead>": "</h5>",
    "<subhead>": (
        '<h6 class="Heading_heading__okScq Heading_head...ing__Ecn_I heading_sm__u3F2n" '
        'data-testid="infoblock-heading">'
    ),
    "</subhead>": "</h6>",
    "<howto_facts>": (
        '<p class="Paragraph_paragraph__exhQA Paragraph_paragraph--default-sm-default__jy0uG '
        'articleParagraph">'
    ),
    "</howto_facts>": "</p><br>",
}


def cueprint_txt_to_cueweb_html(content: str) -> str:
    """Zet Cue Print-code (txt) om naar Cue Web HTML-code.

    Args:
        content: De volledige tekst uit het geüploade .txt-bestand.

    Returns:
        De geconverteerde HTML-code als string.
    """
    # Primaire vervangingen
    for old, new in _REPLACEMENTS.items():
        content = content.replace(old, new)

    # Extra witruimte: wanneer een facts-paragraaf direct gevolgd wordt door een headline,
    # voeg één extra <br> toe.
    # Robust patroon: sluiting van paragraaf + <br> gevolgd door een h5 met data-testid infoblock-headline.
    content = re.sub(
        r"(</p><br>\s*)(<h5\b[^>]*data-testid=\"infoblock-headline\"[^>]*>)",
        r"\1<br>\2",
        content,
        flags=re.IGNORECASE,
    )

    return content
