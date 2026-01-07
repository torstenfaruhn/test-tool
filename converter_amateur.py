import io
import pandas as pd
import numpy as np

# ----------------------------
# Helpers
# ----------------------------

def to_clean_str(x):
    if pd.isna(x):
        return ""
    return str(x).strip()

def parse_int_safe(s):
    try:
        if s is None:
            return None
        s = str(s).strip()
        if s == "":
            return None
        return int(float(s.replace(",", ".")))
    except Exception:
        return None

# ----------------------------
# AANGEPAST: alleen INVOER-tabblad inlezen
# ----------------------------
def load_all_sheets(filebytes: bytes) -> pd.DataFrame:
    """Lees alleen tabblad 'INVOER' (geen FORMULE meer)."""
    xls = pd.ExcelFile(io.BytesIO(filebytes))
    frames = []
    for sheet in xls.sheet_names:
        if sheet != "INVOER":
            continue
        df = pd.read_excel(io.BytesIO(filebytes), sheet_name=sheet, header=0)
        df["__sheet__"] = sheet
        frames.append(df)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

# ----------------------------
# Vind doelpuntenmakerskolom
# ----------------------------
def find_scorers_column(df: pd.DataFrame):
    candidates = [
        c for c in df.columns
        if isinstance(c, str)
        and any(k in c.lower() for k in ["doelpunt", "makers", "scorer"])
    ]
    if candidates:
        return df[candidates[0]].apply(to_clean_str)

    best_i, best_score = None, -1
    for i, c in enumerate(df.columns):
        if i <= 10:
            continue
        s = df[c]
        cnt = 0
        for val in s.dropna().astype(str).values[:500]:
            try:
                float(val.replace(",", "."))
            except:
                cnt += 1
        if cnt > best_score:
            best_score, best_i = cnt, i

    if best_i is not None:
        return df.iloc[:, best_i].apply(to_clean_str)
    return pd.Series([""] * len(df))

# ----------------------------
# AANGEPAST: BEKER wordt divisie
# ----------------------------
def looks_like_division(text: str) -> bool:
    t = str(text or "").strip().lower()
    return (
        ("divisie" in t)
        or ("klasse" in t)
        or ("beker" in t)
    )

# ----------------------------
# Hoofdconversie
# ----------------------------
def excel_to_txt_amateur(file_bytes: bytes) -> str:
    raw = load_all_sheets(file_bytes)
    if raw.empty:
        raise RuntimeError("Geen data gevonden in het Excelbestand.")

    def get_col(df, idx, fallback):
        if df.shape[1] > idx:
            return df.iloc[:, idx].apply(to_clean_str)
        return pd.Series([""] * len(df), name=fallback)

    # vaste kolommen (INVOER)
    home = get_col(raw, 1, "Thuisclub")
    away = get_col(raw, 3, "Uitclub")
    hg   = get_col(raw, 5, "ThuisGoals")
    ag   = get_col(raw, 7, "UitGoals")
    hht  = get_col(raw, 8, "RustThuis")
    aht  = get_col(raw, 10, "RustUit")
    scor = find_scorers_column(raw)

    lines = ["<body>"]

    second_col_header = str(raw.columns[1]) if len(raw.columns) > 1 else ""
    current_div = second_col_header.upper() if looks_like_division(second_col_header) else None
    emitted_div = False

    for i in range(len(raw)):
        home_cell = home.iloc[i]
        away_cell = away.iloc[i]
        hg_raw = hg.iloc[i]
        ag_raw = ag.iloc[i]
        hht_raw = hht.iloc[i]
        aht_raw = aht.iloc[i]
        scorers = scor.iloc[i] if i < len(scor) else ""

        if looks_like_division(home_cell):
            current_div = home_cell.upper()
            emitted_div = False
            continue

        if not (home_cell and home_cell.strip()) or not (away_cell and away_cell.strip()):
            continue

        if current_div and not emitted_div:
            lines.append(f"<subhead_lead>{current_div}</subhead_lead>")
            emitted_div = True

        hg_text = to_clean_str(hg_raw)
        postponed = ("afg" in hg_text.lower()) or ("gest" in hg_text.lower())
        hg_num = parse_int_safe(hg_raw)
        ag_num = parse_int_safe(ag_raw)

        if not postponed and hg_num == 0 and ag_num == 0:
            scorers = " "

        if postponed:
            subhead = f"<subhead>{home_cell} - {away_cell} {hg_text}</subhead>"
        else:
            tg = 0 if hg_num is None else int(hg_num)
            ug = 0 if ag_num is None else int(ag_num)
            rth = 0 if parse_int_safe(hht_raw) is None else int(parse_int_safe(hht_raw))
            rut = 0 if parse_int_safe(aht_raw) is None else int(parse_int_safe(aht_raw))
            subhead = f"<subhead>{home_cell} - {away_cell} {tg}-{ug} ({rth}-{rut})</subhead>"

        lines.append(subhead)
        lines.append("<howto_facts>")
        lines.append(scorers)
        lines.append("</howto_facts>")

    lines.append("</body>")
    return "\n".join(lines)
