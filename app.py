from __future__ import annotations

import os
import uuid
from pathlib import Path
from typing import List

from dotenv import load_dotenv
from flask import Flask, flash, redirect, render_template, request, send_file, url_for
from werkzeug.utils import secure_filename

from agent import apply_business_rules, extract_from_images
from storage import append_json_log, append_to_excel, write_xml_from_excel, read_all_rows_from_excel

load_dotenv()

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
UPLOAD_FOLDER = Path("/tmp/uploads")
DATA_DIR = BASE_DIR / "data"
MASTER_XLSX = BASE_DIR / os.environ.get("MASTER_XLSX", "data/YEIDA_MASTER.xlsx")
MASTER_XML = BASE_DIR / os.environ.get("MASTER_XML", "data/YEIDA_MASTER.xml")
JSON_LOG = DATA_DIR / "YEIDA_MASTER.json"
ALLOWED_EXTENSIONS = {"png", "jpg", "jpeg", "webp", "heic", "heif"}

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "change-this-secret-key")
app.config["MAX_CONTENT_LENGTH"] = int(os.environ.get("MAX_CONTENT_LENGTH_MB", "50")) * 1024 * 1024


def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def save_uploaded_images(files) -> List[Path]:
    batch_id = uuid.uuid4().hex[:12]
    batch_dir = Path("/tmp/uploads") / batch_id
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


@app.route("/", methods=["GET"])
def index():
    rows = read_all_rows_from_excel(MASTER_XLSX)
    return render_template(
        "index.html",
        total_entries=len(rows),
        default_account_type=os.environ.get("DEFAULT_ACCOUNT_TYPE", "31"),
        master_xlsx_exists=MASTER_XLSX.exists(),
        master_xml_exists=MASTER_XML.exists(),
    )


@app.route("/process", methods=["POST"])
def process():
    if "OPENAI_API_KEY" not in os.environ or not os.environ.get("OPENAI_API_KEY"):
        flash("OPENAI_API_KEY missing. Add it in your .env file first.", "error")
        return redirect(url_for("index"))

    files = request.files.getlist("images")
    image_paths = save_uploaded_images(files)
    if not image_paths:
        flash("Please upload at least one image: jpg, jpeg, png, webp, heic, or heif.", "error")
        return redirect(url_for("index"))

    account_type_override = request.form.get("account_type") or None
    category_override = request.form.get("category") or None
    receipt_override = request.form.get("receipt_no") or None
    sol_id_override = request.form.get("sol_id") or None
    user_notes = request.form.get("notes") or None

    try:
        extracted = extract_from_images(image_paths)
        row = apply_business_rules(
            extracted,
            account_type_override=account_type_override,
            category_override=category_override,
            receipt_override=receipt_override,
            sol_id_override=sol_id_override,
        )
        if user_notes:
            row["notes"] = f"{row.get('notes') or ''} {user_notes}".strip()

        new_row_number = append_to_excel(row, MASTER_XLSX)
        write_xml_from_excel(MASTER_XLSX, MASTER_XML)
        append_json_log(row, JSON_LOG)

        return render_template("result.html", row=row, row_number=new_row_number)

    except Exception as exc:
        flash(f"Processing failed: {exc}", "error")
        return redirect(url_for("index"))


@app.route("/download/excel", methods=["GET"])
def download_excel():
    if not MASTER_XLSX.exists():
        flash("Excel file does not exist yet. Upload images first.", "error")
        return redirect(url_for("index"))
    return send_file(MASTER_XLSX, as_attachment=True, download_name="YEIDA_MASTER.xlsx")


@app.route("/download/xml", methods=["GET"])
def download_xml():
    if not MASTER_XML.exists():
        flash("XML file does not exist yet. Upload images first.", "error")
        return redirect(url_for("index"))
    return send_file(MASTER_XML, as_attachment=True, download_name="YEIDA_MASTER.xml")


@app.route("/entries", methods=["GET"])
def entries():
    rows = read_all_rows_from_excel(MASTER_XLSX)
    return render_template("entries.html", rows=rows)


if __name__ == "__main__":
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    app.run(host="0.0.0.0", port=5000, debug=True)
