import io
import re
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict, OrderedDict
from typing import Dict, List, Optional, Set, Tuple


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


# Clubs in deze lijst worden volledig overgeslagen in de output.
# Laat de lijst leeg als je niets wilt uitsluiten.
# Clubnamen worden bij vergelijking hoofdletterongevoelig behandeld
# en dubbele spaties maken niet uit.
EXCLUDED_CLUBS: Tuple[str, ...] = ()


def clean_whitespace(text: str) -> str:
    text = str(text or "")
    text = text.replace("\r", "\n").replace("\xa0", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r" *\n *", "\n", text)
    return text.strip()


def normalize_header(text: str) -> str:
    """
    Maak een kolomkop geschikt voor betrouwbare vergelijking.

    Excel/Forms kan onder meer non-breaking spaces, regeleinden en
    verschillende hoofdletters gebruiken. Die verschillen zijn voor
    kolomherkenning niet relevant.
    """
    return clean_whitespace(text).casefold()


def normalize_club_for_exclude(club_name: str) -> str:
    return clean_whitespace(club_name).casefold()


def build_exclude_set(extra_exclude_clubs: Optional[List[str]] = None) -> Set[str]:
    clubs = list(EXCLUDED_CLUBS)

    if extra_exclude_clubs:
        clubs.extend(extra_exclude_clubs)

    return {
        normalized
        for normalized in (normalize_club_for_exclude(club) for club in clubs)
        if normalized
    }


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

    tokens = [
        strip_trailing_periods(token.strip())
        for token in entry.split(",")
        if token.strip()
    ]
    if not tokens:
        return []

    if (
        len(tokens) >= 3
        and tokens[-1].lower().startswith("allen ")
        and all(len(token.split()) >= 2 for token in tokens[:-1])
    ):
        club = normalize_country_parens(tokens[-1].strip("() "))
        return [
            f"{strip_trailing_periods(name)} ({club})"
            for name in tokens[:-1]
        ]

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
    """
    Lees het eerste werkblad rechtstreeks uit de .xlsx-container.

    De functie retourneert per rijnummer een dictionary:
    {"A": "waarde", "B": "waarde", ...}.
    """
    try:
        workbook = zipfile.ZipFile(io.BytesIO(file_bytes))
    except zipfile.BadZipFile as exc:
        raise RuntimeError(
            "Kon Excelbestand niet openen. Upload een geldig .xlsx-bestand."
        ) from exc

    try:
        shared_strings = _read_shared_strings(workbook)

        if "xl/worksheets/sheet1.xml" not in workbook.namelist():
            raise RuntimeError(
                "Het eerste werkblad ontbreekt in het Excelbestand."
            )

        sheet_root = ET.fromstring(
            workbook.read("xl/worksheets/sheet1.xml")
        )
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

            if (
                cell_type == "s"
                and value_node is not None
                and value_node.text is not None
            ):
                shared_index = int(value_node.text)
                if 0 <= shared_index < len(shared_strings):
                    value = shared_strings[shared_index]
                else:
                    value = ""
            elif cell_type == "inlineStr" and inline_node is not None:
                value = "".join(
                    node.text or ""
                    for node in inline_node.iterfind(".//a:t", NS)
                )
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


def column_sort_key(column: str) -> int:
    """Zet Excel-kolomletters om naar een getal voor sortering."""
    number = 0
    for char in column:
        number = number * 26 + ord(char) - 64
    return number


def find_exact_header(
    headers: Dict[str, str],
    expected_header: str,
) -> str:
    expected = normalize_header(expected_header)

    matches = [
        column
        for column, value in headers.items()
        if normalize_header(value) == expected
    ]

    if len(matches) == 1:
        return matches[0]

    if not matches:
        raise RuntimeError(
            f"Verplichte kolom ontbreekt: '{expected_header}'. "
            "Controleer of dit het juiste Forms/Excel-bestand is."
        )

    raise RuntimeError(
        f"Kolomkop '{expected_header}' komt meerdere keren voor. "
        "De converter kan daardoor niet veilig bepalen welke kolom bedoeld is."
    )


def find_player_columns(
    headers: Dict[str, str],
    regular_prefix: str,
    extra_prefix: str,
    description: str,
) -> List[str]:
    """
    Zoek alle spelerkolommen aan de hand van hun kolomkop.

    Hiermee is de converter niet meer afhankelijk van vaste letters zoals
    T:AF of AG:AS. Een extra of ontbrekende metadata-kolom vóór de
    spelersvelden verschuift de gegevens dan niet meer.
    """
    regular = normalize_header(regular_prefix)
    extra = normalize_header(extra_prefix)

    columns = []

    for column, value in headers.items():
        header = normalize_header(value)

        if header.startswith(regular) or header.startswith(extra):
            columns.append(column)

    columns.sort(key=column_sort_key)

    if not columns:
        raise RuntimeError(
            f"Geen kolommen voor {description} gevonden. "
            "Controleer of de kolomkoppen van het Excelbestand nog "
            "overeenkomen met het formulier."
        )

    return columns


def detect_column_layout(
    rows: Dict[int, Dict[str, str]]
) -> Tuple[str, str, str, List[str], List[str]]:
    """
    Bepaal de benodigde bronkolommen op basis van de kopteksten in rij 1.

    De oude converter gebruikte vaste kolomletters. Microsoft Forms/Excel
    kan echter metadata-kolommen toevoegen of weglaten, waardoor alle
    kolommen daarna opschuiven. Door de kopteksten te gebruiken blijft de
    mapping correct.
    """
    headers = rows.get(1, {})
    if not headers:
        raise RuntimeError(
            "De kopregel in rij 1 ontbreekt in het Excelbestand."
        )

    club_column = find_exact_header(headers, "Naam vereniging")
    division_column = find_exact_header(headers, "Divisie of klasse")
    trainer_column = find_exact_header(headers, "Naam hoofdtrainer")

    new_players_columns = find_player_columns(
        headers,
        regular_prefix="Nieuwe speler (voornaam + achternaam, club)",
        extra_prefix="Indien nog meer nieuwe spelers",
        description="nieuwe spelers",
    )

    departed_players_columns = find_player_columns(
        headers,
        regular_prefix="Vertrokken speler (voornaam + achternaam, club)",
        extra_prefix="Indien nog meer vertrokken spelers",
        description="vertrokken spelers",
    )

    return (
        club_column,
        division_column,
        trainer_column,
        new_players_columns,
        departed_players_columns,
    )


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

    match = re.match(
        r"^(eerste|tweede|derde|vierde|vijfde)\s+klasse\s+([a-z])$",
        normalized,
    )
    if match:
        return (
            2,
            RANK_ORDER[match.group(1)],
            match.group(2),
        )

    return (2, 99, normalized)


def excel_to_txt_mutaties(
    file_bytes: bytes,
    exclude_clubs: Optional[List[str]] = None,
) -> str:
    rows = load_first_sheet_rows(file_bytes)

    (
        club_column,
        division_column,
        trainer_column,
        new_players_columns,
        departed_players_columns,
    ) = detect_column_layout(rows)

    exclude_set = build_exclude_set(exclude_clubs)
    items = []

    for row_number in sorted(rows):
        if row_number == 1:
            continue

        row = rows[row_number]

        club = clean_whitespace(row.get(club_column, ""))
        division = clean_whitespace(row.get(division_column, ""))

        if not club or not division:
            continue

        if normalize_club_for_exclude(club) in exclude_set:
            continue

        trainer = strip_trailing_periods(
            clean_whitespace(row.get(trainer_column, ""))
        )

        nieuwe_spelers = join_player_fields(
            [row.get(column, "") for column in new_players_columns]
        )

        vertrokken_spelers = join_player_fields(
            [row.get(column, "") for column in departed_players_columns]
        )

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
        raise RuntimeError(
            "Geen verwerkbare mutaties gevonden in het Excelbestand."
        )

    groups = defaultdict(list)
    original_labels = OrderedDict()

    for item in items:
        key = normalize_class_for_matching(item["division"])
        groups[key].append(item)
        original_labels.setdefault(key, item["division"])

    ordered_keys = sorted(
        groups.keys(),
        key=lambda key: class_sort_key(original_labels[key]),
    )

    lines = ["<body>"]

    for key in ordered_keys:
        label = original_labels[key]

        # Bestaande Cue Print-opzet:
        # <subhead_lead> bevat divisie/klasse.
        # <subhead> bevat de naam van de vereniging.
        lines.append(f"<subhead_lead>{label}</subhead_lead>")

        for index, item in enumerate(groups[key]):
            if index > 0:
                lines.append("<EP,1>")

            lines.append(f"<subhead>{item['club']}</subhead>")
            lines.append(
                "<howto_facts><bold><CO,5>Nieuw: </bold>"
                f"{item['nieuwe_spelers']}</howto_facts>"
            )
            lines.append(
                "<howto_facts><bold><CO,5>Vertrokken: </bold>"
                f"{item['vertrokken_spelers']}</howto_facts>"
            )
            lines.append(
                "<howto_facts><bold><CO,5>Trainer: </bold>"
                f"{item['trainer']}</howto_facts>"
            )

    lines.append("</body>")
    return "\n".join(lines)
