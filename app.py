from flask import Flask, render_template, request, Response, abort
from converter_regiosport import excel_to_txt_regiosport
from converter_amateur import excel_to_txt_amateur

# Cue Print -> Cue Web converter (Optie 1: volledige classnamen)
from converter_amateur_online import cueprint_txt_to_cueweb_html
from converter_topscorers import extract_text_from_upload, topscorers_text_to_cueweb_html

from openpyxl import Workbook
import io

app = Flask(__name__)


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
    return Response(
        txt,
        mimetype="text/plain; charset=utf-8",
        headers={"Content-Disposition": "attachment; filename=cue_export_regiosport.txt"},
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
    return Response(
        txt,
        mimetype="text/plain; charset=utf-8",
        headers={"Content-Disposition": "attachment; filename=cue_export_amateur.txt"},
    )


@app.post("/convert/amateur-online")
def convert_amateur_online():
    """Converteer Cue Print-uitvoer (txt) naar Cue Web HTML-code (als tekstbestand)."""
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
        content_out = cueprint_txt_to_cueweb_html(content_in)
    except Exception as e:
        return abort(400, f"Kon Amateurvoetbal online-bestand niet verwerken: {e}")

    # Let op: inhoud is HTML-code, maar we leveren het als .txt (kopieerbaar/plakbaar).
    return Response(
        content_out,
        mimetype="text/plain; charset=utf-8",
        headers={"Content-Disposition": "attachment; filename=cue_web_export_amateur.txt"},
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

    return Response(
        html_out,
        mimetype="text/plain; charset=utf-8",
        headers={"Content-Disposition": "attachment; filename=cue_web_export_topscorers.txt"},
    )


# -----------------------------
# Main (alleen voor lokaal testen)
# -----------------------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8000, debug=False)
