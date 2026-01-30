#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
converter_topscorers.py (Render-compatible build)

Key fix:
- Robust template parsing: supports both escaped HTML (&lt;...&gt;) and real HTML (<...>).
- Exposes `convert_topscorers_upload(file_bytes, filename)` for app import.

Dependencies:
- python-docx (only needed for .docx input)
"""

from __future__ import annotations

from pathlib import Path
import argparse
import html as _html
import re
from typing import List, Tuple, Optional

try:
    from docx import Document  # python-docx
except Exception:
    Document = None  # type: ignore


# ---------------------------------------------------------------------------
# HTML templates (CLEANED PER STRATEGIE 1)
# You may keep these escaped or unescaped; parser now accepts both.
# ---------------------------------------------------------------------------

# Escaped variant (keeps compatibility with existing pipeline)
TEMPLATE_HTML = """&lt;ol data-testid="numbered-list"&gt;

&lt;li&gt;
&lt;p&gt;speler 1&lt;/p&gt;
&lt;/li&gt;

&lt;li&gt;
&lt;p&gt;speler 2&lt;br&gt;
speler 3&lt;/p&gt;
&lt;/li&gt;

&lt;/ol&gt;"""

HEADING_TEMPLATE = (
    '&lt;h4 data-testid="article-subhead"&gt;{title}&lt;/h4&gt;'
)


# ---------------------------------------------------------------------------
# Regex helpers (supports escaped and unescaped HTML)
# ---------------------------------------------------------------------------

# Escaped patterns
ESC_OL_RE = re.compile(r"(&lt;ol[^&gt;]*&gt;)(.*?)(&lt;/ol&gt;)", re.S)
ESC_LI_RE = re.compile(r"(&lt;li\b.*?&lt;/li&gt;)", re.S)
ESC_P_RE  = re.compile(r"(&lt;p\b[^&gt;]*&gt;)(.*?)(&lt;/p&gt;)", re.S)

# Unescaped patterns
RAW_OL_RE = re.compile(r"(<ol[^>]*>)(.*?)(</ol>)", re.S | re.IGNORECASE)
RAW_LI_RE = re.compile(r"(<li\b.*?</li>)", re.S | re.IGNORECASE)
RAW_P_RE  = re.compile(r"(<p\b[^>]*>)(.*?)(</p>)", re.S | re.IGNORECASE)

NUMBER_RE = re.compile(r"^\s*\d+\.\s")
GOALS_RE = re.compile(r"\b\d+\b")
DOELPUNT_RE = re.compile(r"\bdoelpunt", re.IGNORECASE)


def _parse_template_blocks(template_text: str) -> Tuple[str, str, str]:
    """
    Try escaped first, then unescaped. Returns (prefix, item_template, suffix).
    """
    # --- Attempt escaped form ---
    m_ol = ESC_OL_RE.search(template_text)
    if m_ol:
