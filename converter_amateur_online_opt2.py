"""Amateurvoetbal online: Cue Print (txt) -> Cue Web (HTML-code)

Optie 2: minimale markup (robust; geen afhankelijkheid van hashed classnamen).
"""

from __future__ import annotations

import re

_REPLACEMENTS: dict[str, str] = {
    "<body>": "",
    "</body>": "",
    "<subhead_lead>": '<h5 data-testid="infoblock-headline">',
    "</subhead_lead>": "</h5>",
    "<subhead>": '<h6 data-testid="infoblock-heading">',
    "</subhead>": "</h6>",
    "<howto_facts>": '<p class="articleParagraph">',
    "</howto_facts>": "</p><br>",
}


def cueprint_txt_to_cueweb_html(content: str) -> str:
    """Zet Cue Print-code (txt) om naar Cue Web HTML-code."""
    for old, new in _REPLACEMENTS.items():
        content = content.replace(old, new)

    content = re.sub(
        r"(</p><br>\s*)(<h5\b[^>]*data-testid=\"infoblock-headline\"[^>]*>)",
        r"\1<br>\2",
        content,
        flags=re.IGNORECASE,
    )

    return content
