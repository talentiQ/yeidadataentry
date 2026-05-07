from __future__ import annotations

import os
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

# Writable temp folders (important for Vercel/serverless)
UPLOAD_FOLDER = Path("/tmp/uploads")
TEMP_DATA_DIR = Path("/tmp/data")

UPLOAD_FOLDER.mkdir(parents=True, exist_ok=True)
TEMP_DATA_DIR.mkdir(parents=True, exist_ok=True)

# Master files
MASTER_XLSX = TEMP_DATA_DIR / "YEIDA_MASTER.xlsx"
MASTER_XML = TEMP_DATA_DIR / "YEIDA_MASTER.xml"
JSON_LOG = TEMP_DATA_DIR / "YEIDA_MASTER.json"

# Allowed image formats
ALLOWED_EXTENSIONS = {
    "png",
    "jpg",
    "jpeg",
    "webp",
    "heic",
    "heif",
}

# =========================================================
# FLASK APP
# =========================================================

app = Flask(__name__)

app.secret_key = os.environ.get(
    "FLASK_SECRET_KEY",
    "change-this-secret-key"
)

app.config["MAX_CONTENT_LENGTH"] = (
    int(os.environ.get("MAX_CONTENT_LENGTH_MB", "50"))
    * 1024
    * 1024
)

# =========================================================
# HELPERS
# =========================================================

def allowed_file(filename: str) -> bool:
    return (
        "." in filename
        and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS
    )


def save_uploaded_images(files) -> List[Path]:
    batch_id = uuid.uuid4().hex[:12]

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

    return saved_paths


# =========================================================
# ROUTES
# =========================================================

@app.route("/", methods=["GET"])
def index():

    rows = []

    if MASTER_XLSX.exists():
        rows = read_all_rows_from_excel(MASTER_XLSX)

    return render_template(
        "index.html",
        total_entries=len(rows),
        default_account_type=os.environ.get(
            "DEFAULT_ACCOUNT_TYPE",
            "31",
        ),
        master_xlsx_exists=MASTER_XLSX.exists(),
        master_xml_exists=MASTER_XML.exists(),
    )


@app.route("/process", methods=["POST"])
def process():

    # API key check
    if (
        "OPENAI_API_KEY" not in os.environ
        or not os.environ.get("OPENAI_API_KEY")
    ):
        flash(
            "OPENAI_API_KEY missing. Add it in your .env file first.",
            "error",
        )
        return redirect(url_for("index"))

    # Uploaded files
    files = request.files.getlist("images")

    image_paths = save_uploaded_images(files)

    if not image_paths:
        flash(
            "Please upload at least one image: jpg, jpeg, png, webp, heic, or heif.",
            "error",
        )
        return redirect(url_for("index"))

    # Optional form overrides — all take priority over OCR extraction
    account_type_override = request.form.get("account_type") or None
    category_override     = request.form.get("category") or None
    receipt_override      = request.form.get("receipt_no") or None
    sol_id_override       = request.form.get("sol_id") or None
    plot_size_override    = request.form.get("plot_size") or None   # FIX: was missing
    user_notes            = request.form.get("notes") or None

    try:

        # OCR extraction
        extracted = extract_from_images(image_paths)

        # Business logic
        row = apply_business_rules(
            extracted,
            account_type_override=account_type_override,
            category_override=category_override,
            receipt_override=receipt_override,
            sol_id_override=sol_id_override,
            plot_size_override=plot_size_override,              # FIX: was missing
        )

        # Append user notes (sol_id is already embedded inside row["notes"]
        # by apply_business_rules; here we only append the free-text notes)
        if user_notes:
            existing = row.get("notes") or ""
            row["notes"] = f"{existing} {user_notes}".strip()

        # Save data
        new_row_number = append_to_excel(row, MASTER_XLSX)

        write_xml_from_excel(
            MASTER_XLSX,
            MASTER_XML,
        )

        append_json_log(
            row,
            JSON_LOG,
        )

        return render_template(
            "result.html",
            row=row,
            row_number=new_row_number,
        )

    except Exception as exc:

        flash(f"Processing failed: {exc}", "error")

        return redirect(url_for("index"))


@app.route("/download/excel", methods=["GET"])
def download_excel():

    if not MASTER_XLSX.exists():
        flash(
            "Excel file does not exist yet. Upload images first.",
            "error",
        )
        return redirect(url_for("index"))

    return send_file(
        MASTER_XLSX,
        as_attachment=True,
        download_name="YEIDA_MASTER.xlsx",
    )


@app.route("/download/xml", methods=["GET"])
def download_xml():

    if not MASTER_XML.exists():
        flash(
            "XML file does not exist yet. Upload images first.",
            "error",
        )
        return redirect(url_for("index"))

    return send_file(
        MASTER_XML,
        as_attachment=True,
        download_name="YEIDA_MASTER.xml",
    )


@app.route("/entries", methods=["GET"])
def entries():

    rows = []

    if MASTER_XLSX.exists():
        rows = read_all_rows_from_excel(MASTER_XLSX)

    return render_template(
        "entries.html",
        rows=rows,
    )


# =========================================================
# MAIN
# =========================================================

if __name__ == "__main__":

    UPLOAD_FOLDER.mkdir(
        parents=True,
        exist_ok=True,
    )

    TEMP_DATA_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    app.run(
        host="0.0.0.0",
        port=5000,
        debug=True,
    )