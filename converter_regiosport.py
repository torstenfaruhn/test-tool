import io
import re
import unicodedata
import pandas as pd


def _strip_accents(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))


def _nl_sort_key(sport: str):
    s = (sport or "").strip()
    if not s:
        return (True, "~")
    s_norm = _strip_accents(s).lower()
    if s_norm.startswith("ij"):
        s_norm = "y" + s_norm[2:]
    return (False, s_norm)


def convert_sheet1_blocks(df):
    """Parse 'Sporten met uitslagregel' naar blokken met keys: sport, render_lines(list).
       Neemt dynamisch alle 'UITSLAGREGEL N' mee (N = 1..∞)."""
    label_col = df.columns[0]
    value_col = df.columns[1]
    blocks = []
    current = {"SPORT": None, "EVENEMENT": None, "UITSLAGREGELS": []}
    uireg = re.compile(r"^UITSLAGREGEL\s*(\d+)$", re.IGNORECASE)

    def flush():
        nonlocal current
        if current["SPORT"] or current["EVENEMENT"] or current["UITSLAGREGELS"]:
            lines = []
            if current.get("SPORT"):
                lines.append(f"<subhead_lead>{current['SPORT']}</subhead_lead><EP>")
            if current.get("EVENEMENT"):
                lines.append(f"<subhead>{current['EVENEMENT']}</subhead><EP>")
            for txt in current["UITSLAGREGELS"]:
                if txt:
                    lines.append(f"<howto_facts>{txt}</howto_facts><EP>")
            blocks.append({"sport": (current.get("SPORT") or "").strip(), "render_lines": lines})
        current = {"SPORT": None, "EVENEMENT": None, "UITSLAGREGELS": []}

    for _, row in df.iterrows():
        label = (str(row.get(label_col)).strip() if pd.notna(row.get(label_col)) else "")
        value = (str(row.get(value_col)).strip() if pd.notna(row.get(value_col)) else "")

        if not label and not value:
            flush(); continue
        if label.upper().startswith("INVOERVELD"):
            flush(); continue

        lab_up = label.upper()
        if lab_up == "SPORT":
            if value: current["SPORT"] = value
        elif lab_up == "EVENEMENT":
            if value: current["EVENEMENT"] = value
        elif uireg.match(lab_up):
            if value: current["UITSLAGREGELS"].append(value)
    flush()
    return blocks


def iter_sheet2_blocks(df):
    """Yield blokken uit 'Sporten met stand' met: sport, evenement, rows, stand."""
    cols = list(df.columns)
    a, b, c, d, e = cols[0], cols[1], cols[2], cols[3], cols[4]
    i, n = 0, len(df)
    while i < n:
        label = str(df.at[i, a]).strip() if pd.notna(df.at[i, a]) else ""
        if label == "SPORT":
            sport = str(df.at[i, b]).strip() if pd.notna(df.at[i, b]) else ""
            i += 1
            evenement = ""
            if i < n and str(df.at[i, a]).strip() == "EVENEMENT":
                evenement = str(df.at[i, b]).strip() if pd.notna(df.at[i, b]) else ""
                i += 1
            # Header overslaan
            if i < n and pd.isna(df.at[i, a]) and all(pd.notna(df.at[i, col]) for col in [b, c, d, e]):
                i += 1
            rows = []
            stand_text = ""
            while i < n:
                lab = str(df.at[i, a]).strip() if pd.notna(df.at[i, a]) else ""
                if lab == "STAND":
                    stand_text = str(df.at[i, b]).strip() if pd.notna(df.at[i, b]) else ""
                    i += 1
                    break
                if lab.startswith("INVOERVELD") or lab == "SPORT":
                    break
                home = str(df.at[i, b]).strip() if pd.notna(df.at[i, b]) else ""
                hs   = str(df.at[i, c]).strip() if pd.notna(df.at[i, c]) else ""
                away = str(df.at[i, d]).strip() if pd.notna(df.at[i, d]) else ""
                ascr = str(df.at[i, e]).strip() if pd.notna(df.at[i, e]) else ""
                if any([home, hs, away, ascr]):
                    rows.append((home, hs, away, ascr))
                i += 1
            yield {"sport": sport, "evenement": evenement, "rows": rows, "stand": stand_text}
        else:
            i += 1


def render_table_block(block):
    lines = []
    lines.append(f"<subhead_lead>{block['sport']}</subhead_lead><EP>")
    lines.append(f"<subhead>{block['evenement']}</subhead><EP>")
    lines.append('<TABLE cciformat="1,0" cols="4" dispwidth="30m" topgutter="0.5272m" bottomgutter="0.5272m" break="norule">')
    lines.extend(['<TCOL width="40m">','</TCOL>','<TCOL width="4m">','</TCOL>','<TCOL width="2m" align="center">','</TCOL>',
                  '<TCOL width="4m" align="right" raster="uniform" color="3,2" pagespot="0" pattern="0" tint="100" angle="0" frequency="0">','</TCOL>'])
    lines.append('<TBODY>')
    n = len(block["rows"])
    for idx, (home, hs, away, ascr) in enumerate(block["rows"]):
        attrs = []
        if idx == 0:     attrs.append('topgutter="1.5816m"')
        if idx == n - 1: attrs.append('bottomgutter="1.5816m"')
        attr_str = f" {' '.join(attrs)}" if attrs else ""
        lines.append(f"<TROW{attr_str}>")
        
# --- Uitzonderingsregel: uitslagen 'n.n.b.', 'afgelast', 'gestaakt' ---
        special_vals = {"n.n.b.", "afgelast", "gestaakt"}
        if hs.lower() in special_vals or ascr.lower() in special_vals:
            # Neem de ingevulde speciale waarde over (behoud oorspronkelijke schrijfwijze)
            special = hs if hs.lower() in special_vals else ascr
            lines += [
                "<TFIELD>", f"{home} - {away}", "</TFIELD>",
                f"<TFIELD colspan='3' align='right'>{special}</TFIELD>"
            ]
        else:
            lines += [
                "<TFIELD>", f"{home} - {away}", "</TFIELD>",
                "<TFIELD>", f"{hs}", "</TFIELD>",
                "<TFIELD>", "-", "</TFIELD>",
                "<TFIELD>", f"{ascr}", "</TFIELD>"
            ]
        lines.append("</TROW>")
    lines.append("</TBODY>")
    lines.append("</TABLE>")
    if block.get("stand"):
        lines.append(f"<howto_facts>{block['stand']}</howto_facts><EP>")
    return lines



def to_render_blocks(sheet1_df, sheet2_df):
    blocks_s1 = convert_sheet1_blocks(sheet1_df)
    blocks_s2 = []
    for b in iter_sheet2_blocks(sheet2_df):
        if not (b['sport'] or b['evenement'] or b['rows']):
            continue
        blocks_s2.append({"sport": b['sport'], "render_lines": render_table_block(b)})
    all_blocks = blocks_s1 + blocks_s2
    return sorted(all_blocks, key=lambda bl: _nl_sort_key(bl.get("sport")))


def suppress_redundant_sportheads(blocks):
    out = []
    last_sport_norm = None
    for bl in blocks:
        sport_norm = _strip_accents((bl.get("sport") or "").strip()).lower()
        if sport_norm.startswith("ij"):
            sport_norm = "y" + sport_norm[2:]
        lines = list(bl["render_lines"])
        if last_sport_norm is not None and sport_norm == last_sport_norm:
            if lines and lines[0].startswith("<subhead_lead>"):
                lines = lines[1:]
        else:
            last_sport_norm = sport_norm
        out.append({"sport": bl.get("sport",""), "render_lines": lines})
    return out


def excel_to_txt_regiosport(file_bytes: bytes) -> str:
    buf = io.BytesIO(file_bytes)
    xls = pd.ExcelFile(buf, engine="openpyxl")
    sheet1 = pd.read_excel(xls, sheet_name=0, dtype=str)
    sheet2 = pd.read_excel(xls, sheet_name=1, dtype=str)

    blocks = to_render_blocks(sheet1, sheet2)
    blocks = suppress_redundant_sportheads(blocks)

    lines = ["<body>"]
    for bl in blocks:
        lines += bl["render_lines"]
    lines.append("</body>")
    output_text = "\n".join(lines)

    # Nabehandeling: <howto_facts> gevolgd door <subhead> krijgt EP,1
    output_text = re.sub(r'</howto_facts><EP>\s*<subhead>', r'</howto_facts><EP,1>\n<subhead>', output_text)

    return output_text
