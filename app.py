from __future__ import annotations

import os
import shutil
import uuid
from pathlib import Path
from typing import List, Tuple

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

BASE_DIR = Path(__file__).resolve().parent

UPLOAD_FOLDER = Path("/tmp/uploads")
TEMP_DATA_DIR = Path("/tmp/data")

UPLOAD_FOLDER.mkdir(parents=True, exist_ok=True)
TEMP_DATA_DIR.mkdir(parents=True, exist_ok=True)

MASTER_XLSX = TEMP_DATA_DIR / "YEIDA_MASTER.xlsx"
MASTER_XML  = TEMP_DATA_DIR / "YEIDA_MASTER.xml"
JSON_LOG    = TEMP_DATA_DIR / "YEIDA_MASTER.json"

ALLOWED_IMAGE_EXTENSIONS = {"png", "jpg", "jpeg", "webp", "heic", "heif"}
ALLOWED_EXTENSIONS = ALLOWED_IMAGE_EXTENSIONS | {"pdf"}

# PDF render resolution
PDF_DPI = 150
PDF_MAX_PAGES = 20

# PyMuPDF — imported at module level so Pylance can resolve it.
# If not installed, PDF uploads will raise a clear RuntimeError at request time.
try:
    import fitz  # type: ignore[import]
    _FITZ_AVAILABLE = True
except ImportError:
    fitz = None  # type: ignore[assignment]
    _FITZ_AVAILABLE = False

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "change-this-secret-key")
app.config["MAX_CONTENT_LENGTH"] = (
    int(os.environ.get("MAX_CONTENT_LENGTH_MB", "50")) * 1024 * 1024
)


def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def is_pdf(filename: str) -> bool:
    return filename.rsplit(".", 1)[-1].lower() == "pdf"


def pdf_to_images(pdf_path: Path, output_dir: Path) -> List[Path]:
    """Convert each page of a scanned PDF to PNG using PyMuPDF."""
    if not _FITZ_AVAILABLE or fitz is None:
        raise RuntimeError(
            "PyMuPDF is not installed. Run: pip install pymupdf"
        )

    mat = fitz.Matrix(PDF_DPI / 72, PDF_DPI / 72)
    image_paths: List[Path] = []

    with fitz.open(pdf_path) as doc:
        total_pages = min(len(doc), PDF_MAX_PAGES)
        for page_num in range(total_pages):
            pix = doc[page_num].get_pixmap(matrix=mat, alpha=False)
            img_path = output_dir / f"page_{page_num + 1:03d}.png"
            pix.save(str(img_path))
            image_paths.append(img_path)

    return image_paths


def save_uploaded_files(files) -> Tuple[List[Path], Path]:
    """
    Save uploaded images or PDFs. PDFs are converted to per-page PNGs.
    Returns (image_paths, batch_dir). Caller must clean up batch_dir.
    """
    batch_dir = UPLOAD_FOLDER / uuid.uuid4().hex[:12]
    batch_dir.mkdir(parents=True, exist_ok=True)
    image_paths: List[Path] = []

    for file in files:
        if not file or not file.filename:
            continue
        if not allowed_file(file.filename):
            continue

        filename = secure_filename(file.filename)
        saved_path = batch_dir / filename
        file.save(saved_path)

        if is_pdf(filename):
            pdf_images = pdf_to_images(saved_path, batch_dir)
            image_paths.extend(pdf_images)
            saved_path.unlink(missing_ok=True)  # free space after conversion
        else:
            image_paths.append(saved_path)

    return image_paths, batch_dir


@app.route("/", methods=["GET"])
def index():
    rows = read_all_rows_from_excel(MASTER_XLSX) if MASTER_XLSX.exists() else []
    return render_template(
        "index.html",
        total_entries=len(rows),
        default_account_type=os.environ.get("DEFAULT_ACCOUNT_TYPE", "31"),
        master_xlsx_exists=MASTER_XLSX.exists(),
        master_xml_exists=MASTER_XML.exists(),
        upload_accept=".png,.jpg,.jpeg,.webp,.heic,.heif,.pdf",
    )


@app.route("/process", methods=["POST"])
def process():

    if (
        "OPENAI_API_KEY" not in os.environ
        or not os.environ.get("OPENAI_API_KEY")
    ):
        flash("OPENAI_API_KEY missing. Add it in your .env file first.", "error")
        return redirect(url_for("index"))

    files = request.files.getlist("images")
    batch_dir = None
    image_paths, batch_dir = save_uploaded_files(files)

    if not image_paths:
        flash(
            "Please upload at least one image or PDF. "
            "Accepted: jpg, jpeg, png, webp, heic, heif, pdf.",
            "error",
        )
        if batch_dir:
            shutil.rmtree(batch_dir, ignore_errors=True)
        return redirect(url_for("index"))

    account_type_override = request.form.get("account_type") or None
    category_override     = request.form.get("category") or None
    receipt_override      = request.form.get("receipt_no") or None
    sol_id_override       = request.form.get("sol_id") or None
    user_notes            = request.form.get("notes") or None

    plot_size_override = request.form.get("plot_size") or None
    try:
        plot_size_override = int(plot_size_override) if plot_size_override else None
    except (ValueError, TypeError):
        plot_size_override = None

    try:
        extracted = extract_from_images(image_paths)

        row = apply_business_rules(
            extracted,
            account_type_override=account_type_override,
            category_override=category_override,
            receipt_override=receipt_override,
            sol_id_override=sol_id_override,
            plot_size_override=plot_size_override,
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

    finally:
        if batch_dir:
            shutil.rmtree(batch_dir, ignore_errors=True)


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


@app.route("/download/json", methods=["GET"])
def download_json():
    if not JSON_LOG.exists():
        flash("JSON log does not exist yet. Upload images first.", "error")
        return redirect(url_for("index"))
    return send_file(JSON_LOG, as_attachment=True, download_name="YEIDA_MASTER.json")


@app.route("/entries", methods=["GET"])
def entries():
    rows = read_all_rows_from_excel(MASTER_XLSX) if MASTER_XLSX.exists() else []
    return render_template("entries.html", rows=rows)


if __name__ == "__main__":
    UPLOAD_FOLDER.mkdir(parents=True, exist_ok=True)
    TEMP_DATA_DIR.mkdir(parents=True, exist_ok=True)
    app.run(host="0.0.0.0", port=5000, debug=True)