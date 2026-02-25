from flask import Flask, render_template, request, Response, abort, jsonify, make_response
from converter_regiosport import excel_to_txt_regiosport
from converter_amateur import excel_to_txt_amateur

# Cue Print -> Cue Web converter
from converter_amateur_online import cueprint_txt_to_docx_bytes
from converter_topscorers import extract_text_from_upload_bytes, topscorers_text_to_docx_bytes
from converter_topscorers_cumulated import cumulated_topscorers_to_docx_bytes, ConversionError

import os
import tempfile
import secrets
import shutil
from pathlib import Path

from openpyxl import Workbook
import io

from datetime import datetime
from zoneinfo import ZoneInfo
from urllib.parse import quote
import re
import time

app = Flask(__name__)


@app.after_request
def _security_headers(resp):
    # Geen tracking / third-party. Alleen eigen scripts en styles.
    resp.headers.setdefault(
        "Content-Security-Policy",
        "default-src 'self'; img-src 'self' data:; style-src 'self'; script-src 'self'; base-uri 'none'; object-src 'none'; frame-ancestors 'none'; form-action 'self'",
    )
    resp.headers.setdefault("Referrer-Policy", "no-referrer")
    resp.headers.setdefault("X-Content-Type-Options", "nosniff")
    resp.headers.setdefault("X-Frame-Options", "DENY")
    return resp


# -----------------------------
# Output bestandsnamen
# -----------------------------
#
# Placeholders:
# - {date}: YYYYMMDD (Europe/Amsterdam)
# - {date_dash}: YYYY-MM-DD (Europe/Amsterdam)
# - {time}: HHMM (24u, Europe/Amsterdam)
# - {stem}: uploadnaam zonder extensie (geschoond)
#
AMATEUR_OUTPUT_PATTERN = "{date}_cue_print_uitslagen_amateurs.txt"
AMATEUR_ONLINE_OUTPUT_PATTERN = "{date}_cue_word_uitslagen_amateurs.docx"
REGIOSPORT_OUTPUT_PATTERN = "{date}_cue_print_uitslagen_regiosport.txt"
TOPSCORERS_OUTPUT_PATTERN = "{date}_cue_word_topscorers_amateurs.docx"
TOPSCORERS_CUMULATED_OUTPUT_PATTERN = "{date}_cue_word_gecumuleerde_topscorers_amateurs.docx"


def _sanitize_stem(filename: str) -> str:
    """Maak een veilige, korte bestands-stem op basis van de uploadnaam."""
    name = (filename or "").strip()
    name = name.split("/")[-1].split("\\")[-1]
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
# Tijdelijke sessie-opslag (in /tmp)
# - Uploads worden opgeslagen per sessie-cookie.
# - Bij 'Maak gecumuleerde Word-export' wordt altijd direct opgeruimd (succes én fout).
# -----------------------------
SESSION_COOKIE = "dlst"
MAX_UPLOAD_BYTES = 15 * 1024 * 1024  # 15 MB

_last_cleanup_ts = 0.0

def _maybe_cleanup_tmp_sessions() -> None:
    global _last_cleanup_ts
    now = time.time()
    # Max 1x per uur per worker opruimen
    if now - _last_cleanup_ts < 3600:
        return
    _last_cleanup_ts = now

    base = tempfile.gettempdir()
    try:
        for name in os.listdir(base):
            if not name.startswith("dlst_"):
                continue
            p = os.path.join(base, name)
            try:
                st = os.stat(p)
            except Exception:
                continue
            # ouder dan 6 uur -> weg
            if (now - st.st_mtime) > 6 * 3600:
                try:
                    shutil.rmtree(p, ignore_errors=True)
                except Exception:
                    pass
    except Exception:
        pass


def _get_session_token() -> str | None:
    tok = request.cookies.get(SESSION_COOKIE)
    if not tok:
        return None
    if not re.fullmatch(r"[A-Za-z0-9_-]{20,200}", tok):
        return None
    return tok


def _new_session_token() -> str:
    return secrets.token_urlsafe(32)


def _session_dir(token: str) -> str:
    return os.path.join(tempfile.gettempdir(), f"dlst_{token}")


def _ensure_session_dir(token: str) -> str:
    p = _session_dir(token)
    os.makedirs(p, exist_ok=True)
    return p


def _save_upload(token: str, kind: str, raw: bytes, filename: str) -> None:
    if len(raw) > MAX_UPLOAD_BYTES:
        raise ValueError("Bestand is te groot.")
    d = _ensure_session_dir(token)
    with open(os.path.join(d, f"{kind}.bin"), "wb") as f:
        f.write(raw)
    with open(os.path.join(d, f"{kind}.name.txt"), "w", encoding="utf-8") as f:
        f.write(filename or "")


def _load_upload(token: str, kind: str) -> tuple[bytes, str]:
    d = _session_dir(token)
    with open(os.path.join(d, f"{kind}.bin"), "rb") as f:
        raw = f.read()
    try:
        name = Path(os.path.join(d, f"{kind}.name.txt")).read_text(encoding="utf-8")
    except Exception:
        name = ""
    return raw, name


def _clear_session_dir(token: str | None) -> None:
    if not token:
        return
    d = _session_dir(token)
    try:
        shutil.rmtree(d, ignore_errors=True)
    except Exception:
        pass


def _clear_session_cookie(resp: Response) -> Response:
    resp.set_cookie(SESSION_COOKIE, "", max_age=0, path="/", samesite="Lax", httponly=True)
    return resp


@app.get("/")
def index():
    return render_template("index.html")


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
    """Converteer Cue Print-uitvoer (txt met tags) naar Word (.docx)."""
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
        docx_bytes, _stats = cueprint_txt_to_docx_bytes(content_in)
    except Exception as e:
        return abort(400, f"Kon Amateurvoetbal online-bestand niet verwerken: {e}")

    out_name = _build_output_filename(AMATEUR_ONLINE_OUTPUT_PATTERN, file.filename or "")

    return Response(
        docx_bytes,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": _content_disposition_attachment(out_name)},
    )


def _xls_bytes_from_workbook(wb: Workbook) -> bytes:
    """Helper: zet Workbook om naar bytes voor download."""
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.getvalue()


@app.post("/convert/topscorers")
def convert_topscorers():
    """Converteer topscorers-tekst (.txt/.docx) naar Word (.docx) voor Cue Web."""
    file = request.files.get("file_topscorers")
    if not file or file.filename == "":
        return abort(400, "Geen bestand geüpload (Topscorers).")

    if (file.filename or "").lower().endswith(".doc"):
        return abort(400, "Verkeerd bestandstype: .doc wordt niet ondersteund. Sla op als .docx of .txt.")

    try:
        raw = file.read()

        # prevent obvious wrong uploads (.xlsx / zip)
        if raw.startswith(b"PK\x03\x04") and not (file.filename or "").lower().endswith(".docx"):
            return abort(
                400,
                "Verkeerd bestand: dit lijkt geen .txt of .docx. Upload een tekstbestand (Word of Kladblok).",
            )

        text_in = extract_text_from_upload_bytes(raw, file.filename or "")
        docx_bytes = topscorers_text_to_docx_bytes(text_in)
    except Exception as e:
        return abort(400, f"Kon topscorers-bestand niet verwerken: {e}")

    out_name = _build_output_filename(TOPSCORERS_OUTPUT_PATTERN, file.filename or "")

    return Response(
        docx_bytes,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": _content_disposition_attachment(out_name)},
    )


@app.post("/upload/topscorers-cumulated/source")
def upload_topscorers_cumulated_source():
    _maybe_cleanup_tmp_sessions()
    file = request.files.get("file_source")
    if not file or file.filename == "":
        return jsonify({"ok": False, "code": "TS-CUM-001", "message": "Geen bronbestand geüpload."}), 400

    fn = file.filename.lower()
    if not (fn.endswith(".docx") or fn.endswith(".doc") or fn.endswith(".txt")):
        return jsonify({"ok": False, "code": "TS-CUM-002", "message": "Verkeerd bestandstype. Upload .doc, .docx of .txt."}), 400

    token = _get_session_token() or _new_session_token()

    try:
        _save_upload(token, "source", file.read(), file.filename)
    except Exception:
        return jsonify({"ok": False, "code": "TS-CUM-008", "message": "Kon bronbestand niet opslaan."}), 400

    resp = make_response(jsonify({"ok": True}))
    resp.set_cookie(SESSION_COOKIE, token, httponly=True, samesite="Lax", path="/")
    return resp


@app.post("/upload/topscorers-cumulated/results")
def upload_topscorers_cumulated_results():
    _maybe_cleanup_tmp_sessions()
    file = request.files.get("file_results")
    if not file or file.filename == "":
        return jsonify({"ok": False, "code": "TS-CUM-001", "message": "Geen uitslagenbestand geüpload."}), 400

    fn = file.filename.lower()
    if not (fn.endswith(".xlsx") or fn.endswith(".xls")):
        return jsonify({"ok": False, "code": "TS-CUM-002", "message": "Verkeerd bestandstype. Upload .xlsx of .xls."}), 400

    token = _get_session_token() or _new_session_token()

    try:
        _save_upload(token, "results", file.read(), file.filename)
    except Exception:
        return jsonify({"ok": False, "code": "TS-CUM-008", "message": "Kon uitslagenbestand niet opslaan."}), 400

    resp = make_response(jsonify({"ok": True}))
    resp.set_cookie(SESSION_COOKIE, token, httponly=True, samesite="Lax", path="/")
    return resp


@app.post("/convert/topscorers-cumulated")
def convert_topscorers_cumulated():
    """Maak gecumuleerde topscorers Word-export (bron-stand + Excel-ronde)."""
    token = _get_session_token()
    if not token:
        resp = Response("TS-CUM-001: Upload eerst beide bestanden.", status=400, mimetype="text/plain; charset=utf-8")
        return _clear_session_cookie(resp)

    try:
        source_raw, source_name = _load_upload(token, "source")
        results_raw, results_name = _load_upload(token, "results")
    except Exception:
        _clear_session_dir(token)
        resp = Response("TS-CUM-001: Upload eerst beide bestanden.", status=400, mimetype="text/plain; charset=utf-8")
        return _clear_session_cookie(resp)

    try:
        docx_bytes = cumulated_topscorers_to_docx_bytes(source_raw, source_name, results_raw, results_name)
        out_name = _build_output_filename(TOPSCORERS_CUMULATED_OUTPUT_PATTERN, source_name or "topscorers")
        resp = Response(
            docx_bytes,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": _content_disposition_attachment(out_name)},
        )
        return _clear_session_cookie(resp)
    except ConversionError as e:
        resp = Response(str(e), status=400, mimetype="text/plain; charset=utf-8")
        return _clear_session_cookie(resp)
    except Exception:
        resp = Response("TS-CUM-007: Onverwachte fout tijdens verwerken.", status=400, mimetype="text/plain; charset=utf-8")
        return _clear_session_cookie(resp)
    finally:
        _clear_session_dir(token)



if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8000, debug=False)
