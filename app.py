from __future__ import annotations

import os
import shutil
<<<<<<< HEAD
import threading
import uuid
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Any, Dict, List, Tuple
=======
import uuid
from pathlib import Path
from typing import List, Tuple
>>>>>>> 1ff615e8134835d15bece98a02955ed2f494de09

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

<<<<<<< HEAD
# =========================================================
# BASE PATHS
# =========================================================

=======
>>>>>>> 1ff615e8134835d15bece98a02955ed2f494de09
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

<<<<<<< HEAD
# PDF SETTINGS
PDF_DPI          = 100
PDF_MAX_PAGES    = 25
PDF_JPEG_QUALITY = 75

# =========================================================
# THREADING CONFIG
# =========================================================

# How many candidates are OCR-processed simultaneously.
# Each worker = one concurrent OpenAI API call.
# Recommended: 3-5 on localhost (limited by your OpenAI rate limit tier).
# Tier 1 (default): ~500 RPM → safe at 5 workers.
# Tier 2+: can go higher.
MAX_OCR_WORKERS = int(os.environ.get("MAX_OCR_WORKERS", "4"))

# Global lock — ensures only ONE thread writes to Excel/JSON at a time.
# OCR (OpenAI call) runs in parallel. Only the save step is serialised.
_WRITE_LOCK = threading.Lock()

# =========================================================
# FLASK APP
# =========================================================

# PyMuPDF
=======
# PDF render resolution
PDF_DPI = 150
PDF_MAX_PAGES = 20

# PyMuPDF — imported at module level so Pylance can resolve it.
# If not installed, PDF uploads will raise a clear RuntimeError at request time.
>>>>>>> 1ff615e8134835d15bece98a02955ed2f494de09
try:
    import fitz  # type: ignore[import]
    _FITZ_AVAILABLE = True
except ImportError:
    fitz = None  # type: ignore[assignment]
    _FITZ_AVAILABLE = False

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "change-this-secret-key")
<<<<<<< HEAD

app.config["MAX_CONTENT_LENGTH"] = (
    int(os.environ.get("MAX_CONTENT_LENGTH_MB", "1000")) * 1024 * 1024
)

# =========================================================
# FILE HELPERS
# =========================================================
=======
app.config["MAX_CONTENT_LENGTH"] = (
    int(os.environ.get("MAX_CONTENT_LENGTH_MB", "50")) * 1024 * 1024
)

>>>>>>> 1ff615e8134835d15bece98a02955ed2f494de09

def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def is_pdf(filename: str) -> bool:
    return filename.rsplit(".", 1)[-1].lower() == "pdf"


def pdf_to_images(pdf_path: Path, output_dir: Path) -> List[Path]:
<<<<<<< HEAD
    """Convert PDF pages to compressed JPEGs for OCR."""
    if not _FITZ_AVAILABLE or fitz is None:
        raise RuntimeError("PyMuPDF not installed. Run: pip install pymupdf")
=======
    """Convert each page of a scanned PDF to PNG using PyMuPDF."""
    if not _FITZ_AVAILABLE or fitz is None:
        raise RuntimeError(
            "PyMuPDF is not installed. Run: pip install pymupdf"
        )
>>>>>>> 1ff615e8134835d15bece98a02955ed2f494de09

    mat = fitz.Matrix(PDF_DPI / 72, PDF_DPI / 72)
    image_paths: List[Path] = []

    with fitz.open(pdf_path) as doc:
<<<<<<< HEAD
        for page_num in range(min(len(doc), PDF_MAX_PAGES)):
            pix      = doc[page_num].get_pixmap(matrix=mat, alpha=False)
            img_path = output_dir / f"{pdf_path.stem}_page_{page_num + 1:03d}.jpg"
            pix.save(str(img_path), output="jpeg", jpg_quality=PDF_JPEG_QUALITY)
=======
        total_pages = min(len(doc), PDF_MAX_PAGES)
        for page_num in range(total_pages):
            pix = doc[page_num].get_pixmap(matrix=mat, alpha=False)
            img_path = output_dir / f"page_{page_num + 1:03d}.png"
            pix.save(str(img_path))
>>>>>>> 1ff615e8134835d15bece98a02955ed2f494de09
            image_paths.append(img_path)

    return image_paths


<<<<<<< HEAD
def compress_image(src: Path, max_width: int = 1800) -> Path:
    """Resize + re-compress an image to reduce payload size."""
    try:
        from PIL import Image as PilImage
        img = PilImage.open(src)
        if img.mode not in ("RGB", "L"):
            img = img.convert("RGB")
        w, h = img.size
        if w > max_width:
            img = img.resize((max_width, int(h * max_width / w)), PilImage.Resampling.LANCZOS)
        dst = src.with_suffix(".jpg")
        img.save(dst, "JPEG", quality=80, optimize=True)
        if dst != src:
            src.unlink(missing_ok=True)
        return dst
    except Exception:
        return src


def save_uploaded_files(files) -> Tuple[List[Dict[str, Any]], Path]:
    """
    Save uploaded files and group them applicant-wise.
    Each file (PDF or image) = one applicant.

    Returns: (applicants_list, batch_dir)
      applicants_list = [{"source": filename, "images": [Path, ...]}, ...]
    """
    batch_dir = UPLOAD_FOLDER / uuid.uuid4().hex[:12]
    batch_dir.mkdir(parents=True, exist_ok=True)
    applicants: List[Dict[str, Any]] = []
=======
def save_uploaded_files(files) -> Tuple[List[Path], Path]:
    """
    Save uploaded images or PDFs. PDFs are converted to per-page PNGs.
    Returns (image_paths, batch_dir). Caller must clean up batch_dir.
    """
    batch_dir = UPLOAD_FOLDER / uuid.uuid4().hex[:12]
    batch_dir.mkdir(parents=True, exist_ok=True)
    image_paths: List[Path] = []
>>>>>>> 1ff615e8134835d15bece98a02955ed2f494de09

    for file in files:
        if not file or not file.filename:
            continue
        if not allowed_file(file.filename):
            continue

<<<<<<< HEAD
        filename   = secure_filename(file.filename)
=======
        filename = secure_filename(file.filename)
>>>>>>> 1ff615e8134835d15bece98a02955ed2f494de09
        saved_path = batch_dir / filename
        file.save(saved_path)

        if is_pdf(filename):
            pdf_images = pdf_to_images(saved_path, batch_dir)
<<<<<<< HEAD
            applicants.append({"source": filename, "images": pdf_images})
            saved_path.unlink(missing_ok=True)
        else:
            compressed = compress_image(saved_path)
            applicants.append({"source": filename, "images": [compressed]})

    return applicants, batch_dir


# =========================================================
# PROCESS ONE CANDIDATE
# =========================================================

def _process_one_candidate(
    applicant: Dict[str, Any],
    overrides: Dict[str, Any],
) -> Dict[str, Any]:
    """
    Process a single candidate: OCR → business rules → save to Excel.

    This function is called from multiple threads simultaneously.
    The OpenAI OCR call is fully parallel.
    The Excel/JSON write is protected by _WRITE_LOCK so only one
    thread touches the file at a time — prevents corruption.

    Returns a result dict:
      success → {"source": ..., "ok": True,  "row_number": N, "row": {...}}
      failure → {"source": ..., "ok": False, "error": "..."}
    """
    source = applicant["source"]

    try:
        # ── PARALLEL: OpenAI OCR call ─────────────────────────────────
        extracted = extract_from_images(applicant["images"])

        # ── PARALLEL: business rules (pure CPU, no I/O) ───────────────
        row = apply_business_rules(
            extracted,
            account_type_override = overrides.get("account_type") or None,
            category_override     = overrides.get("category") or None,
            receipt_override      = overrides.get("receipt_no") or None,
            sol_id_override       = overrides.get("sol_id") or None,
            plot_size_override    = overrides.get("plot_size_int"),
        )

        if overrides.get("notes"):
            row["notes"] = f"{row.get('notes') or ''} {overrides['notes']}".strip()

        # ── SERIALISED: Excel + JSON write ────────────────────────────
        # Lock ensures only one thread writes at a time.
        # OCR for other candidates continues running while this thread
        # holds the lock — the wait is only ~0.5s per write.
        with _WRITE_LOCK:
            row_number = append_to_excel(row, MASTER_XLSX)
            append_json_log(row, JSON_LOG)

        return {"source": source, "ok": True, "row_number": row_number, "row": row}

    except Exception as exc:
        return {"source": source, "ok": False, "error": str(exc)}


# =========================================================
# ROUTES
# =========================================================

=======
            image_paths.extend(pdf_images)
            saved_path.unlink(missing_ok=True)  # free space after conversion
        else:
            image_paths.append(saved_path)

    return image_paths, batch_dir


>>>>>>> 1ff615e8134835d15bece98a02955ed2f494de09
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
<<<<<<< HEAD
        max_workers=MAX_OCR_WORKERS,
=======
>>>>>>> 1ff615e8134835d15bece98a02955ed2f494de09
    )


@app.route("/process", methods=["POST"])
def process():

<<<<<<< HEAD
    if not os.environ.get("OPENAI_API_KEY"):
        flash("OPENAI_API_KEY missing. Add it in your .env file first.", "error")
        return redirect(url_for("index"))

    files     = request.files.getlist("images")
    batch_dir = None

    applicants, batch_dir = save_uploaded_files(files)

    if not applicants:
=======
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
>>>>>>> 1ff615e8134835d15bece98a02955ed2f494de09
        flash(
            "Please upload at least one image or PDF. "
            "Accepted: jpg, jpeg, png, webp, heic, heif, pdf.",
            "error",
        )
        if batch_dir:
            shutil.rmtree(batch_dir, ignore_errors=True)
        return redirect(url_for("index"))

<<<<<<< HEAD
    # Build overrides dict — passed to every worker thread
    plot_size_raw = request.form.get("plot_size") or None
    try:
        plot_size_int = int(plot_size_raw) if plot_size_raw else None
    except (ValueError, TypeError):
        plot_size_int = None

    overrides = {
        "account_type": request.form.get("account_type") or None,
        "category":     request.form.get("category") or None,
        "receipt_no":   request.form.get("receipt_no") or None,
        "sol_id":       request.form.get("sol_id") or None,
        "notes":        request.form.get("notes") or None,
        "plot_size_int": plot_size_int,
    }

    success_results: List[Dict] = []
    failed_files:    List[Dict] = []

    try:
        # ── Parallel OCR — all candidates processed simultaneously ────
        # Workers = min(MAX_OCR_WORKERS, number of candidates)
        workers = min(MAX_OCR_WORKERS, len(applicants))

        with ThreadPoolExecutor(max_workers=workers) as executor:
            futures = {
                executor.submit(_process_one_candidate, applicant, overrides): applicant["source"]
                for applicant in applicants
            }

            for future in as_completed(futures):
                result = future.result()
                if result["ok"]:
                    success_results.append(result)
                else:
                    failed_files.append({
                        "file":  result["source"],
                        "error": result["error"],
                    })
                    print(f"[FAILED] {result['source']}: {result['error']}")

        # ── Rebuild XML once after all rows are written ───────────────
        if success_results:
            write_xml_from_excel(MASTER_XLSX, MASTER_XML)

        rows = read_all_rows_from_excel(MASTER_XLSX) if MASTER_XLSX.exists() else []

        return render_template(
            "result.html",
            success_count=len(success_results),
            failed_count=len(failed_files),
            failed_files=failed_files,
            total_entries=len(rows),
        )

=======
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

>>>>>>> 1ff615e8134835d15bece98a02955ed2f494de09
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


<<<<<<< HEAD
# =========================================================
# MAIN
# =========================================================

if __name__ == "__main__":
    UPLOAD_FOLDER.mkdir(parents=True, exist_ok=True)
    TEMP_DATA_DIR.mkdir(parents=True, exist_ok=True)
    # threaded=True is required — Flask must handle concurrent requests
    # from the thread pool while the main request is still open.
    app.run(host="0.0.0.0", port=5000, debug=True, threaded=True)
=======
if __name__ == "__main__":
    UPLOAD_FOLDER.mkdir(parents=True, exist_ok=True)
    TEMP_DATA_DIR.mkdir(parents=True, exist_ok=True)
    app.run(host="0.0.0.0", port=5000, debug=True)
>>>>>>> 1ff615e8134835d15bece98a02955ed2f494de09
