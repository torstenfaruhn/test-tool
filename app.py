from flask import Flask, render_template, request, Response, abort, send_file
from converter_regiosport import excel_to_txt_regiosport
from converter_amateur import excel_to_txt_amateur

# Nieuw: Cue Print -> Cue Web (HTML) converter
from converter_amateur_online import cueprint_txt_to_cueweb_html

# Voor sjablonen
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
    """Converteer Cue Print-uitvoer (txt) naar Cue Web HTML-code."""
    file = request.files.get("file_amateur_online")
    if not file or file.filename == "":
        return abort(400, "Geen bestand geüpload (Amateurvoetbal online).")
    try:
        content_in = file.read().decode("utf-8", errors="replace")
        content_out = cueprint_txt_to_cueweb_html(content_in)
    except Exception as e:
        return abort(400, f"Kon Amateurvoetbal online-bestand niet verwerken: {e}")
    return Response(
        content_out,
        mimetype="text/plain; charset=utf-8",
        headers={"Content-Disposition": "attachment; filename=cue_web_export_amateur.html"},
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




# -----------------------------
# Main
# -----------------------------
if __name__ == "__main__":
    # Voor lokaal testen:
    # python app.py -> http://localhost:8000
    app.run(host="0.0.0.0", port=8000, debug=False)
