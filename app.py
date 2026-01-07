from __future__ import annotations

from datetime import datetime
from zoneinfo import ZoneInfo

from flask import Flask, Response, abort, render_template, request

from converter_regiosport import excel_to_txt_regiosport
from converter_amateur import excel_to_txt_amateur
from converter_amateur_online import cueprint_txt_to_cueweb_html

# Topscorers (docx/txt -> Cue Web HTML-code as text)
from converter_topscorers import convert_topscorers_upload


app = Flask(__name__)


def yyyymmdd_ams() -> str:
    """Return current date in Europe/Amsterdam as YYYYMMDD."""
    return datetime.now(ZoneInfo("Europe/Amsterdam")).strftime("%Y%m%d")


@app.get("/")
def index():
    return render_template("index.html")


# -----------------------------
# Regiosport: Excel -> CUE txt
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

    fn = f'{yyyymmdd_ams()}_cue_print_uitslagen_regiosport.txt'
    return Response(
        txt,
        mimetype="text/plain; charset=utf-8",
        headers={"Content-Disposition": f'attachment; filename="{fn}"'},
    )


# -----------------------------
# Amateurvoetbal: Excel -> CUE txt
# -----------------------------
@app.post("/convert/amateur")
def convert_amateur():
    file = request.files.get("file_amateur")
    if not file or file.filename == "":
        return abort(400, "Geen bestand geüpload (Amateurvoetbal).")

    try:
        txt = excel_to_txt_amateur(file.read())
    except Exception as e:
        return abort(400, f"Kon Amateur-bestand niet verwerken: {e}")

    fn = f'{yyyymmdd_ams()}_cue_print_uitslagen_amateurvoetbal.txt'
    return Response(
        txt,
        mimetype="text/plain; charset=utf-8",
        headers={"Content-Disposition": f'attachment; filename="{fn}"'},
    )


# -----------------------------
# Amateurvoetbal online: Cue Print txt -> Cue Web HTML-code (as txt)
# -----------------------------
@app.post("/convert/amateur-online")
def convert_amateur_online():
    file = request.files.get("file_amateur_online")
    if not file or file.filename == "":
        return abort(400, "Geen bestand geüpload (Amateurvoetbal online).")

    try:
        raw = file.read()

        # .xlsx (en veel ZIP-bestanden) beginnen met PK\x03\x04; voorkom onbruikbare output.
        # Let op: .docx is hier niet toegestaan (alleen .txt).
        if raw.startswith(b"PK\x03\x04"):
            return abort(
                400,
                "Verkeerd bestand: dit lijkt een Excelbestand (.xlsx). Upload een Cue Print-tekstbestand (.txt).",
            )

        content_in = raw.decode("utf-8", errors="replace")
        content_out = cueprint_txt_to_cueweb_html(content_in)
    except Exception as e:
        return abort(400, f"Kon Amateurvoetbal online-bestand niet verwerken: {e}")

    fn = f'{yyyymmdd_ams()}_cue_web_html_uitslagen_amateurvoetbal.txt'
    return Response(
        content_out,
        mimetype="text/plain; charset=utf-8",
        headers={"Content-Disposition": f'attachment; filename="{fn}"'},
    )


# -----------------------------
# Topscorers: .docx/.txt -> Cue Web HTML-code (as txt)
# -----------------------------
@app.post("/convert/topscorers")
def convert_topscorers():
    file = request.files.get("file_topscorers")
    if not file or file.filename == "":
        return abort(400, "Geen bestand geüpload (Topscorers).")

    filename = file.filename or ""
    raw = file.read()

    # Hier is PK\x03\x04 NIET per definitie fout, want .docx is een ZIP-container.
    # Validatie doen we op extensie; de converter kan met beide overweg.
    name = filename.lower()
    if not (name.endswith(".txt") or name.endswith(".docx")):
        return abort(400, "Verkeerd bestandstype. Upload een .txt of .docx bestand.")

    try:
        content_out = convert_topscorers_upload(raw, filename)
    except Exception as e:
        return abort(400, f"Kon topscorers-bestand niet verwerken: {e}")

    fn = f'{yyyymmdd_ams()}_cue_web_html_cumulatieve_lijst_topscorers.txt'
    return Response(
        content_out,
        mimetype="text/plain; charset=utf-8",
        headers={"Content-Disposition": f'attachment; filename="{fn}"'},
    )


if __name__ == "__main__":
    # Lokaal testen: Render gebruikt gunicorn.
    app.run(host="0.0.0.0", port=8000, debug=False)
