from flask import Flask, render_template, request, Response, abort, send_file, after_this_request
from converter_regiosport import excel_to_txt_regiosport
from converter_amateur import excel_to_txt_amateur

# Cue Print -> DOCX converter (Amateur online)
from converter_amateur_online import cueprint_txt_to_docx

from converter_topscorers import extract_text_from_upload, topscorers_text_to_cueweb_html

from openpyxl import Workbook
import io

from datetime import datetime
from zoneinfo import ZoneInfo
from urllib.parse import quote
import re
import os
import tempfile

app = Flask(__name__)


# -----------------------------
# Output bestandsnamen
# -----------------------------
#
# Pas deze patterns aan om de vaste naamopbouw per converter te bepalen.
# Beschikbare placeholders:
# - {date}: YYYYMMDD (Europe/Amsterdam)
# - {date_dash}: YYYY-MM-DD (Europe/Amsterdam)
# - {time}: HHMM (24u, Europe/Amsterdam)
# - {stem}: bestandsnaam van de upload zónder extensie (geschoond)
#
# Gevraagd vaste formats:
# - converter_amateur          -> YYYYMMDD_cue_print_uitslagen_amateurs.txt
# - converter_amateur_online   -> YYYYMMDD_cue_web_uitslagen_amateurs.docx
# - converter_regiosport       -> YYYYMMDD_cue_print_uitslagen_regiosport.txt
# - converter_topscorers       -> YYYYMMDD_cue_web_topscorers_amateurs.txt

AMATEUR_OUTPUT_PATTERN = "{date}_cue_print_uitslagen_amateurs.txt"
AMATEUR_ONLINE_OUTPUT_PATTERN = "{date}_cue_web_uitslagen_amateurs.docx"
REGIOSPORT_OUTPUT_PATTERN = "{date}_cue_print_uitslagen_regiosport.txt"
TOPSCORERS_OUTPUT_PATTERN = "{date}_cue_web_topscorers_amateurs.txt"


def _sanitize_stem(filename: str) -> str:
    """Maak een veilige, korte bestands-stem op basis van de uploadnaam."""
    name = (filename or "").strip()
    # haal pad-separators weg
    name = name.split("/")[-1].split("\\")[-1]
    # drop extensie
    if "." in name:
        name = name.rsplit(".", 1)[0]

    name = name.strip().replace(" ", "_")
    name = re.sub(r"[^A-Za-z0-9_-]+", "", name)
    name = re.sub(r"_+", "_", name)
    name = name.strip("_-")
    return (name or "input")[:60]


def _content_disposition_attachment(filename: str) -> str:
    """RFC 6266: stuur zowel filename als filename* voor brede browsercompat."""
    safe_ascii = filename.encode("ascii", "ignore").decode("ascii") or "export.txt"
    return f'attachment; filename="{safe_ascii}"; filename*=UTF-8\'\'{quote(filename)}'


def _build_output_filename(pattern: str, uploaded_filename: str) -> str:
    now = datetime.now(ZoneInfo("Europe/Amsterdam"))
    return pattern.format(
        date=now.strftime("%Y%m%d"),
        date_dash=now.strftime("%Y-%m-%d"),
        time=now.strftime("%H%M"),
        stem=_sanitize_stem(uploaded_filename),
    )


# -----------------------------
# UI
# -----------------------------
@app.get("/")
def index():
    return render_template("index.html")


# -----------------------------
# Convert endpoints
# -----------------------------
@app.post("/convert/regiosport")
def convert_regiosport():
    file = request.files.get("file_regio")
    if not file or file.filename == "":
        return abort(400, "Geen bestand geüpload (Regiosport).")
    try:
        txt = excel_to_txt_regiosport(file.read())
    except Exception as e:
        return abort(400, f"Kon Regiosport-bestand niet verwerken: {e}")

    out_name = _build_output_filename(REGIOSPORT_OUTPUT_PATTERN, file.filename or "")
    return Response(
        txt,
        mimetype="text/plain; charset=utf-8",
        headers={"Content-Disposition": _content_disposition_attachment(out_name)},
    )


@app.post("/convert/amateur")
def convert_amateur():
    file = request.files.get("file_amateur")
    if not file or file.filename == "":
        return abort(400, "Geen bestand geüpload (Amateurvoetbal).")
    try:
        txt = excel_to_txt_amateur(file.read())
    except Exception as e:
        return abort(400, f"Kon Amateur-bestand niet verwerken: {e}")

    out_name = _build_output_filename(AMATEUR_OUTPUT_PATTERN, file.filename or "")
    return Response(
        txt,
        mimetype="text/plain; charset=utf-8",
        headers={"Content-Disposition": _content_disposition_attachment(out_name)},
    )


@app.post("/convert/amateur-online")
def convert_amateur_online():
    """Converteer Cue Print-uitvoer (txt) naar DOCX (download)."""
    file = request.files.get("file_amateur_online")
    if not file or file.filename == "":
        return abort(400, "Geen bestand geüpload (Amateurvoetbal online).")

    try:
        raw = file.read()

        # .xlsx (en veel ZIP-bestanden) beginnen met PK\x03\x04; voorkom onbruikbare output.
        if raw.startswith(b"PK\x03\x04"):
            return abort(
                400,
                "Verkeerd bestand: dit lijkt een Excelbestand (.xlsx). Upload een Cue Print-tekstbestand (.txt).",
            )

        content_in = raw.decode("utf-8", errors="replace")

        # Maak tijdelijk docx-bestand
        fd, out_path = tempfile.mkstemp(suffix=".docx", prefix="uitslagen_")
        os.close(fd)

        # Genereer docx
        cueprint_txt_to_docx(content_in, out_path)

    except Exception as e:
        # Probeer tempbestand op te ruimen als het bestaat
        try:
            if "out_path" in locals() and out_path and os.path.exists(out_path):
                os.remove(out_path)
        except OSError:
            pass
        return abort(400, f"Kon Amateurvoetbal online-bestand niet verwerken: {e}")

    out_name = _build_output_filename(AMATEUR_ONLINE_OUTPUT_PATTERN, file.filename or "")

    # Ruim tempbestand op na response
    @after_this_request
    def cleanup(response):
        try:
            if out_path and os.path.exists(out_path):
                os.remove(out_path)
        except OSError:
            pass
        return response

    return send_file(
        out_path,
        as_attachment=True,
        download_name=out_name,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": _content_disposition_attachment(out_name)},
    )


# -----------------------------
# Template (leeg invoerdocument) endpoints
# -----------------------------
def _xls_bytes_from_workbook(wb: Workbook) -> bytes:
    """Helper: zet Workbook om naar bytes voor download."""
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.getvalue()


@app.post("/convert/topscorers")
def convert_topscorers():
    """Converteer topscorers-tekst (.txt/.docx) naar Cue Web HTML-code (als .txt voor copy/paste)."""
    file = request.files.get("file_topscorers")
    if not file or file.filename == "":
        return abort(400, "Geen bestand geüpload (Topscorers).")

    try:
        raw = file.read()

        # prevent obvious wrong uploads (.xlsx / zip)
        if raw.startswith(b"PK\x03\x04") and not (file.filename or "").lower().endswith(".docx"):
            return abort(
                400,
                "Verkeerd bestand: dit lijkt geen .txt of .docx. Upload een tekstbestand (Word of Kladblok).",
            )

        text_in = extract_text_from_upload(raw, file.filename or "")
        html_out = topscorers_text_to_cueweb_html(text_in)
    except Exception as e:
        return abort(400, f"Kon topscorers-bestand niet verwerken: {e}")

    out_name = _build_output_filename(TOPSCORERS_OUTPUT_PATTERN, file.filename or "")

    return Response(
        html_out,
        mimetype="text/plain; charset=utf-8",
        headers={"Content-Disposition": _content_disposition_attachment(out_name)},
    )


# -----------------------------
# Main (alleen voor lokaal testen)
# -----------------------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8000, debug=False)
