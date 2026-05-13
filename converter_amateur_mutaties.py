import io
import re
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict, OrderedDict
from typing import Dict, List


NS = {
    "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
}

RANK_ORDER = {
    "eerste": 1,
    "tweede": 2,
    "derde": 3,
    "vierde": 4,
    "vijfde": 5,
}


# Kolomindeling bron-Excel met extra kolommen P en Q voor assistent-trainer.
CLUB_COLUMN = "G"
DIVISION_COLUMN = "H"
TRAINER_COLUMN = "L"
NEW_PLAYERS_COLUMNS = ("T", "AF")
DEPARTED_PLAYERS_COLUMNS = ("AG", "AS")


def clean_whitespace(text: str) -> str:
    text = str(text or "")
    text = text.replace("\r", "\n").replace("\xa0", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r" *\n *", "\n", text)
    return text.strip()


def strip_trailing_periods(text: str) -> str:
    text = str(text or "")
    text = re.sub(r"\.(?=\s*(?:,|\)|$))", "", text)
    return text.strip()


def normalize_country_parens(text: str) -> str:
    text = str(text or "").strip()
    text = re.sub(r"\s*\(([^()]+)\)\s*$", lambda m: ", " + m.group(1).strip(), text)
    text = re.sub(r"\s*,\s*", ", ", text)
    return text.strip(" ,")


def normalize_existing_parenthetical_entry(entry: str) -> str:
    entry = clean_whitespace(entry)
    match = re.match(r"^(.*?)\s*\((.*)\)\s*$", entry)
    if not match:
        return entry

    name = strip_trailing_periods(clean_whitespace(match.group(1)))
    club = strip_trailing_periods(clean_whitespace(match.group(2)))
    club = normalize_country_parens(club)
    club = strip_trailing_periods(club)
    return f"{name} ({club})"


def normalize_plain_entry(entry: str) -> List[str]:
    entry = clean_whitespace(entry)
    if not entry:
        return []

    entry = re.sub(r"\s+-\s+", ", ", entry)
    entry = re.sub(r"\s*,\s*", ", ", entry)
    entry = strip_trailing_periods(entry)

    if "," not in entry:
        return [normalize_existing_parenthetical_entry(entry)]

    tokens = [strip_trailing_periods(token.strip()) for token in entry.split(",") if token.strip()]
    if not tokens:
        return []

    if len(tokens) >= 3 and tokens[-1].lower().startswith("allen ") and all(len(token.split()) >= 2 for token in tokens[:-1]):
        club = normalize_country_parens(tokens[-1].strip("() "))
        return [f"{strip_trailing_periods(name)} ({club})" for name in tokens[:-1]]

    if len(tokens) == 2:
        name = tokens[0]
        club = tokens[1]
    else:
        if len(tokens[0].split()) >= 2:
            name = tokens[0]
            club = ", ".join(tokens[1:])
        else:
            name = " ".join(tokens[:-1])
            club = tokens[-1]

    name = strip_trailing_periods(clean_whitespace(name))
    club = clean_whitespace(club)

    if re.fullmatch(r"\(.*\)", club):
        club = club[1:-1].strip()

    club = strip_trailing_periods(club)
    club = normalize_country_parens(club)
    club = strip_trailing_periods(club)

    return [f"{name} ({club})"]


def normalize_cell_value(cell_value: str) -> List[str]:
    if cell_value is None:
        return []

    text = clean_whitespace(cell_value)
    if not text:
        return []

    items: List[str] = []
    for part in text.split("\n"):
        part = clean_whitespace(part)
        if not part:
            continue
        items.extend(normalize_plain_entry(part))
    return [item for item in items if item]


def join_player_fields(values: List[str]) -> str:
    items: List[str] = []
    for value in values:
        items.extend(normalize_cell_value(value))
    return ", ".join(items) if items else "niemand"


def col_range(start: str, end: str) -> List[str]:
    def to_num(col: str) -> int:
        number = 0
        for char in col:
            number = number * 26 + ord(char) - 64
        return number

    def to_col(number: int) -> str:
        result = ""
        while number:
            number, remainder = divmod(number - 1, 26)
            result = chr(65 + remainder) + result
        return result

    return [to_col(i) for i in range(to_num(start), to_num(end) + 1)]


def _read_shared_strings(workbook: zipfile.ZipFile) -> List[str]:
    shared_strings: List[str] = []
    if "xl/sharedStrings.xml" not in workbook.namelist():
        return shared_strings

    shared_root = ET.fromstring(workbook.read("xl/sharedStrings.xml"))
    for item in shared_root.findall("a:si", NS):
        parts = [node.text or "" for node in item.iterfind(".//a:t", NS)]
        shared_strings.append("".join(parts))
    return shared_strings


def load_first_sheet_rows(file_bytes: bytes) -> Dict[int, Dict[str, str]]:
    try:
        workbook = zipfile.ZipFile(io.BytesIO(file_bytes))
    except zipfile.BadZipFile as exc:
        raise RuntimeError("Kon Excelbestand niet openen. Upload een geldig .xlsx-bestand.") from exc

    try:
        shared_strings = _read_shared_strings(workbook)
        if "xl/worksheets/sheet1.xml" not in workbook.namelist():
            raise RuntimeError("Het eerste werkblad ontbreekt in het Excelbestand.")

        sheet_root = ET.fromstring(workbook.read("xl/worksheets/sheet1.xml"))
        rows: Dict[int, Dict[str, str]] = defaultdict(dict)

        for cell in sheet_root.findall(".//a:sheetData/a:row/a:c", NS):
            reference = cell.attrib.get("r", "")
            match = re.match(r"([A-Z]+)(\d+)", reference)
            if not match:
                continue

            column = match.group(1)
            row_number = int(match.group(2))
            cell_type = cell.attrib.get("t")
            value_node = cell.find("a:v", NS)
            inline_node = cell.find("a:is", NS)

            if cell_type == "s" and value_node is not None and value_node.text is not None:
                value = shared_strings[int(value_node.text)]
            elif cell_type == "inlineStr" and inline_node is not None:
                value = "".join(node.text or "" for node in inline_node.iterfind(".//a:t", NS))
            elif value_node is not None and value_node.text is not None:
                value = value_node.text
            else:
                value = ""

            rows[row_number][column] = value

        return rows
    except RuntimeError:
        raise
    except Exception as exc:
        raise RuntimeError("Kon het Excelbestand niet lezen.") from exc
    finally:
        workbook.close()


def normalize_class_for_matching(label: str) -> str:
    label = clean_whitespace(label).lower()
    label = label.replace("klassse", "klasse")
    return label


def class_sort_key(label: str) -> tuple:
    normalized = normalize_class_for_matching(label)

    if normalized == "derde divisie b":
        return (0, 0, "")
    if normalized == "vierde divisie c":
        return (1, 0, "")
    if normalized == "vrouwen hoofdklasse":
        return (3, 0, "")
    if normalized == "vrouwen eerste klasse c":
        return (3, 1, "")
    if normalized == "vrouwen eerste klasse d":
        return (3, 2, "")

    match = re.match(r"^(eerste|tweede|derde|vierde|vijfde)\s+klasse\s+([a-z])$", normalized)
    if match:
        return (2, RANK_ORDER[match.group(1)], match.group(2))

    return (2, 99, normalized)


def excel_to_txt_mutaties(file_bytes: bytes) -> str:
    rows = load_first_sheet_rows(file_bytes)
    items = []

    for row_number in sorted(rows):
        if row_number == 1:
            continue

        row = rows[row_number]
        club = clean_whitespace(row.get(CLUB_COLUMN, ""))
        division = clean_whitespace(row.get(DIVISION_COLUMN, ""))
        if not club or not division:
            continue

        trainer = strip_trailing_periods(clean_whitespace(row.get(TRAINER_COLUMN, "")))
        nieuwe_spelers = join_player_fields([row.get(col, "") for col in col_range(*NEW_PLAYERS_COLUMNS)])
        vertrokken_spelers = join_player_fields([row.get(col, "") for col in col_range(*DEPARTED_PLAYERS_COLUMNS)])

        items.append(
            {
                "club": club,
                "division": division,
                "trainer": trainer,
                "nieuwe_spelers": nieuwe_spelers,
                "vertrokken_spelers": vertrokken_spelers,
            }
        )

    if not items:
        raise RuntimeError("Geen verwerkbare mutaties gevonden in het Excelbestand.")

    groups = defaultdict(list)
    original_labels = OrderedDict()

    for item in items:
        key = normalize_class_for_matching(item["division"])
        groups[key].append(item)
        original_labels.setdefault(key, item["division"])

    ordered_keys = sorted(groups.keys(), key=lambda key: class_sort_key(original_labels[key]))

    lines = ["<body>"]

    for key in ordered_keys:
        label = original_labels[key]
        lines.append(f"<subhead_lead>{label}</subhead_lead>")

        for index, item in enumerate(groups[key]):
            if index > 0:
                lines.append("<EP,1>")

            lines.append(f"<subhead>{item['club']}</subhead>")
            lines.append(f"<howto_facts><bold><CO,5>Nieuw: </bold>{item['nieuwe_spelers']}</howto_facts>")
            lines.append(f"<howto_facts><bold><CO,5>Vertrokken: </bold>{item['vertrokken_spelers']}</howto_facts>")
            lines.append(f"<howto_facts><bold><CO,5>Trainer: </bold>{item['trainer']}</howto_facts>")

    lines.append("</body>")
    return "\n".join(lines)
