from __future__ import annotations

import os
import shutil
import uuid
from pathlib import Path
from typing import List

from dotenv import load_dotenv
from flask import (
    Flask,
    flash,
    redirect,
    render_template,
    request,
    send_file,
    url_for,
)
from werkzeug.utils import secure_filename

from agent import apply_business_rules, extract_from_images
from storage import (
    append_json_log,
    append_to_excel,
    write_xml_from_excel,
    read_all_rows_from_excel,
)

load_dotenv()

# =========================================================
# BASE PATHS
# =========================================================

BASE_DIR = Path(__file__).resolve().parent

# Writable temp folders
# NOTE: /tmp is ephemeral on Vercel/serverless — data is lost between invocations.
# For production use a persistent store (S3, Supabase, etc.).
UPLOAD_FOLDER = Path("/tmp/uploads")
TEMP_DATA_DIR = Path("/tmp/data")

UPLOAD_FOLDER.mkdir(parents=True, exist_ok=True)
TEMP_DATA_DIR.mkdir(parents=True, exist_ok=True)

MASTER_XLSX = TEMP_DATA_DIR / "YEIDA_MASTER.xlsx"
MASTER_XML  = TEMP_DATA_DIR / "YEIDA_MASTER.xml"
JSON_LOG    = TEMP_DATA_DIR / "YEIDA_MASTER.json"

ALLOWED_EXTENSIONS = {"png", "jpg", "jpeg", "webp", "heic", "heif"}

# =========================================================
# FLASK APP
# =========================================================

app = Flask(__name__)

app.secret_key = os.environ.get("FLASK_SECRET_KEY", "change-this-secret-key")

app.config["MAX_CONTENT_LENGTH"] = (
    int(os.environ.get("MAX_CONTENT_LENGTH_MB", "50")) * 1024 * 1024
)

# =========================================================
# HELPERS
# =========================================================

def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def save_uploaded_images(files) -> tuple[List[Path], Path]:
    """Save uploaded image files to a unique batch directory.

    Returns (list_of_saved_paths, batch_dir) so the caller can clean up
    the batch directory after processing.
    """
    batch_id  = uuid.uuid4().hex[:12]
    batch_dir = UPLOAD_FOLDER / batch_id
    batch_dir.mkdir(parents=True, exist_ok=True)

    saved_paths = []
    for file in files:
        if not file or not file.filename:
            continue
        if not allowed_file(file.filename):
            continue
        filename = secure_filename(file.filename)
        path = batch_dir / filename
        file.save(path)
        saved_paths.append(path)

    return saved_paths, batch_dir


# =========================================================
# ROUTES
# =========================================================

@app.route("/", methods=["GET"])
def index():
    rows = read_all_rows_from_excel(MASTER_XLSX) if MASTER_XLSX.exists() else []
    return render_template(
        "index.html",
        total_entries=len(rows),
        default_account_type=os.environ.get("DEFAULT_ACCOUNT_TYPE", "31"),
        master_xlsx_exists=MASTER_XLSX.exists(),
        master_xml_exists=MASTER_XML.exists(),
    )


@app.route("/process", methods=["POST"])
def process():
    # ── API key guard ────────────────────────────────────────────────────
    if not os.environ.get("OPENAI_API_KEY"):
        flash(
            "OPENAI_API_KEY missing. Add it to your .env file first.",
            "error",
        )
        return redirect(url_for("index"))

    # ── Uploaded files ───────────────────────────────────────────────────
    files = request.files.getlist("images")

    # batch_dir initialised to None so the finally block never hits a NameError
    # even if save_uploaded_images() itself raises an exception.
    batch_dir = None
    image_paths, batch_dir = save_uploaded_images(files)

    if not image_paths:
        flash(
            "Please upload at least one image: jpg, jpeg, png, webp, heic, or heif.",
            "error",
        )
        shutil.rmtree(batch_dir, ignore_errors=True)
        return redirect(url_for("index"))

    # ── Optional form overrides ──────────────────────────────────────────
    account_type_override = request.form.get("account_type") or None
    category_override     = request.form.get("category") or None
    receipt_override      = request.form.get("receipt_no") or None
    sol_id_override       = request.form.get("sol_id") or None
    user_notes            = request.form.get("notes") or None

    try:
        # OCR extraction
        extracted = extract_from_images(image_paths)

        # Business rules
        row = apply_business_rules(
            extracted,
            account_type_override=account_type_override,
            category_override=category_override,
            receipt_override=receipt_override,
            sol_id_override=sol_id_override,
        )

        if user_notes:
            row["notes"] = f"{row.get('notes') or ''} {user_notes}".strip()

        # Persist
        new_row_number = append_to_excel(row, MASTER_XLSX)
        write_xml_from_excel(MASTER_XLSX, MASTER_XML)
        append_json_log(row, JSON_LOG)

        return render_template("result.html", row=row, row_number=new_row_number)

    except Exception as exc:
        flash(f"Processing failed: {exc}", "error")
        return redirect(url_for("index"))

    finally:
        # Always clean up temp upload batch to prevent disk accumulation.
        # batch_dir is None only if save_uploaded_images() itself threw,
        # in which case there is nothing to clean up.
        if batch_dir:
            shutil.rmtree(batch_dir, ignore_errors=True)


@app.route("/download/excel", methods=["GET"])
def download_excel():
    if not MASTER_XLSX.exists():
        flash("Excel file does not exist yet. Upload images first.", "error")
        return redirect(url_for("index"))
    return send_file(
        MASTER_XLSX,
        as_attachment=True,
        download_name="YEIDA_MASTER.xlsx",
    )


@app.route("/download/xml", methods=["GET"])
def download_xml():
    if not MASTER_XML.exists():
        flash("XML file does not exist yet. Upload images first.", "error")
        return redirect(url_for("index"))
    return send_file(
        MASTER_XML,
        as_attachment=True,
        download_name="YEIDA_MASTER.xml",
    )


@app.route("/download/json", methods=["GET"])
def download_json():
    if not JSON_LOG.exists():
        flash("JSON log does not exist yet. Upload images first.", "error")
        return redirect(url_for("index"))
    return send_file(
        JSON_LOG,
        as_attachment=True,
        download_name="YEIDA_MASTER.json",
    )


@app.route("/entries", methods=["GET"])
def entries():
    rows = read_all_rows_from_excel(MASTER_XLSX) if MASTER_XLSX.exists() else []
    # FIX: read_all_rows_from_excel now returns List[Dict] (not List[List]).
    # entries.html must reference fields by name: row.rps_no, row["rps_no"], etc.
    return render_template("entries.html", rows=rows)


# =========================================================
# MAIN
# =========================================================

if __name__ == "__main__":
    UPLOAD_FOLDER.mkdir(parents=True, exist_ok=True)
    TEMP_DATA_DIR.mkdir(parents=True, exist_ok=True)
    app.run(host="0.0.0.0", port=5000, debug=True)