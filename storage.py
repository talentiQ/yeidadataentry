from __future__ import annotations

import json
import re
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path
from shutil import copyfile
from typing import Dict, List

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


# =========================================================
# TEMPLATE FILE
# =========================================================

BASE_DIR = Path(__file__).resolve().parent

TEMPLATE_FILE = BASE_DIR / "data" / "RPS_10_MASTER.xlsx"

# Last row that has pre-filled validation formulas in the template (AU–BC).
# Detected from the template: formulas exist through row 10050 → index 10050.
TEMPLATE_FORMULA_MAX_ROW = 10050

# Validation formula templates (row placeholder = "{R}").
# These are injected for rows beyond TEMPLATE_FORMULA_MAX_ROW.
_AU = (
    '=IF(N{R}="","",IF(AND('
    'LEN(TRIM(N{R}))=10,'
    'CODE(UPPER(MID(N{R},1,1)))>=65,CODE(UPPER(MID(N{R},1,1)))<=90,'
    'CODE(UPPER(MID(N{R},2,1)))>=65,CODE(UPPER(MID(N{R},2,1)))<=90,'
    'CODE(UPPER(MID(N{R},3,1)))>=65,CODE(UPPER(MID(N{R},3,1)))<=90,'
    'UPPER(MID(N{R},4,1))="P",'
    'CODE(UPPER(MID(N{R},5,1)))>=65,CODE(UPPER(MID(N{R},5,1)))<=90,'
    'ISNUMBER(VALUE(MID(N{R},6,1))),ISNUMBER(VALUE(MID(N{R},7,1))),'
    'ISNUMBER(VALUE(MID(N{R},8,1))),ISNUMBER(VALUE(MID(N{R},9,1))),'
    'CODE(UPPER(MID(N{R},10,1)))>=65,CODE(UPPER(MID(N{R},10,1)))<=90'
    '),"✔ Valid PAN","✘ Invalid PAN"))'
)
_VALIDATION_FORMULAS: Dict[str, str] = {
    "AU": _AU,
    "AV": '=IF(N{R}="","",IF(COUNTIF($N$4:$N$10000,N{R})>1,"✗ DUPLICATE PAN","✓ Unique"))',
    "AW": '=IF(AA{R}="","",IF(COUNTIF($AA$4:$AA$10000,AA{R})>1,"✗ DUPLICATE Mobile","✓ Unique"))',
    "AX": '=IF(AP{R}="","",IF(ISNUMBER(MATCH(VALUE(AP{R}),{162,183,184,200,223,290},0)),"✔ Valid Size","✘ Invalid Size"))',
    "AY": '=IF(M{R}="","",IF(OR(UPPER(TRIM(M{R}))="GEN",UPPER(TRIM(M{R}))="SC/ST",UPPER(TRIM(M{R}))="SC",UPPER(TRIM(M{R}))="ST"),"✔ Valid Cat","✘ Invalid Cat"))',
    "AZ": '=IF(AC{R}="","",IF(UPPER(TRIM(AC{R}))="AUTO DEBIT","✔ Valid IFSC",IF(LEN(TRIM(AF{R}))=11,"✔ Valid IFSC","✘ Invalid IFSC")))',
    "BA": '=IF(AA{R}="","",IF(AND(ISNUMBER(AA{R}),LEN(TEXT(AA{R},"0"))=10),"✔ Valid Mob","✘ Invalid Mob"))',
    "BB": '=IF(D{R}="","",IF(AND(LEN(TRIM(D{R}))=12,EXACT(UPPER(TRIM(D{R})),TRIM(D{R}))),"✔ Valid RPS","✘ RPS Must Be 12 Chars"))',
    "BC": (
        '=IF(AND(N{R}="",AP{R}=""),"",IF('
        'SUMPRODUCT(--(ISNUMBER(SEARCH({"DUPLICATE","ERROR","✗","Invalid"},AU{R}:BB{R}))))>0,'
        '"✗ ERRORS – CHECK",'
        '"✓ READY TO SUBMIT"'
        '))'
    ),
}


# =========================================================
# HELPERS
# =========================================================

def ensure_parent(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)


def _find_next_empty_row(ws) -> int:
    """
    Return the first row (>= 4) where BOTH col D (RPS No) and
    col E (Receipt No) are blank.

    The original code only checked col E and started from row 4,
    which caused it to overwrite row 4 even when row 5 had real data
    because E4 was None (no formula, no value) → while-condition was
    False immediately → next_row stayed at 4.
    """
    row = 4
    while True:
        d_val = ws[f"D{row}"].value
        e_val = ws[f"E{row}"].value
        # A row is "occupied" if either key identifier column has content
        if d_val in (None, "") and e_val in (None, ""):
            return row
        row += 1


def _inject_validation_formulas(ws, row: int) -> None:
    """
    Write validation formulas for columns AU–BC at *row*.

    The template pre-fills these formulas through TEMPLATE_FORMULA_MAX_ROW.
    For rows beyond that limit the formulas are absent, so we inject them
    here to keep validation consistent for every appended row.
    """
    if row <= TEMPLATE_FORMULA_MAX_ROW:
        # Formula already present from the template – nothing to do.
        return
    for col, tmpl in _VALIDATION_FORMULAS.items():
        ws[f"{col}{row}"] = tmpl.replace("{R}", str(row))


# =========================================================
# APPEND TO EXCEL
# =========================================================

def append_to_excel(row: Dict, output_path: Path) -> int:

    ensure_parent(output_path)

    # ── Create output from template the first time ──────────────────────
    if not output_path.exists():
        copyfile(TEMPLATE_FILE, output_path)

    # ── Load WITHOUT data_only so formulas are preserved ────────────────
    # Loading with data_only=True and then saving would permanently strip
    # all formula strings from AU–BC, leaving those cells blank forever.
    wb = load_workbook(output_path)
    ws = wb.active

    # ── Find first truly empty data row ─────────────────────────────────
    next_row = _find_next_empty_row(ws)

    # ── Serial number ────────────────────────────────────────────────────
    ws[f"A{next_row}"] = next_row - 3

    # ── Basic details ────────────────────────────────────────────────────
    ws[f"B{next_row}"] = row.get("lot_no", "")
    ws[f"C{next_row}"] = row.get("file_no", "")
    ws[f"D{next_row}"] = row.get("rps_no", "")
    ws[f"E{next_row}"] = row.get("receipt_no", "")
    ws[f"F{next_row}"] = row.get("sourcing_branch", "")

    # ── Name ─────────────────────────────────────────────────────────────
    ws[f"G{next_row}"] = row.get("first_name", "")
    ws[f"H{next_row}"] = row.get("middle_name", "")
    ws[f"I{next_row}"] = row.get("last_name", "")

    # ── Personal details ─────────────────────────────────────────────────
    ws[f"J{next_row}"] = row.get("dob", "")
    ws[f"K{next_row}"] = row.get("gender", "")
    ws[f"L{next_row}"] = row.get("marital_status", "")
    ws[f"M{next_row}"] = row.get("category", "")
    ws[f"N{next_row}"] = row.get("pan_no", "")
    ws[f"O{next_row}"] = row.get("religion", "")
    ws[f"P{next_row}"] = row.get("current_residence", "")
    ws[f"Q{next_row}"] = row.get("qualification", "")
    ws[f"R{next_row}"] = row.get("profession", "")

    # ── Address ──────────────────────────────────────────────────────────
    ws[f"S{next_row}"] = row.get("address_1", "")
    ws[f"T{next_row}"] = row.get("address_2", "")
    ws[f"U{next_row}"] = row.get("address_3", "")
    ws[f"V{next_row}"] = row.get("city", "")
    ws[f"W{next_row}"] = row.get("state", "")
    ws[f"X{next_row}"] = row.get("zip_code", "")

    # ── Contact ──────────────────────────────────────────────────────────
    ws[f"Y{next_row}"] = row.get("phone_1", "")
    ws[f"Z{next_row}"] = row.get("phone_2", "")
    ws[f"AA{next_row}"] = row.get("mobile", "")
    ws[f"AB{next_row}"] = row.get("email", "")

    # ── Bank details ─────────────────────────────────────────────────────
    ws[f"AC{next_row}"] = row.get("repayment_mode", "")
    ws[f"AD{next_row}"] = row.get("account_no", "")
    ws[f"AE{next_row}"] = row.get("micr", "")
    ws[f"AF{next_row}"] = row.get("ifsc", "")
    ws[f"AG{next_row}"] = row.get("bank_name", "")
    ws[f"AH{next_row}"] = row.get("account_type", "")

    # ── KYC ──────────────────────────────────────────────────────────────
    ws[f"AI{next_row}"] = row.get("id_proof_no", "")
    ws[f"AJ{next_row}"] = row.get("id_expiry", "")
    ws[f"AK{next_row}"] = row.get("id_doc_type", "")
    ws[f"AL{next_row}"] = row.get("address_proof_no", "")
    ws[f"AM{next_row}"] = row.get("address_expiry", "")
    ws[f"AN{next_row}"] = row.get("address_doc_type", "")

    # ── SC/ST ─────────────────────────────────────────────────────────────
    ws[f"AO{next_row}"] = row.get("sc_st_flag", "")

    # ── Property ─────────────────────────────────────────────────────────
    ws[f"AP{next_row}"] = row.get("plot_size", "")
    ws[f"AQ{next_row}"] = row.get("asset_cost", "")
    ws[f"AR{next_row}"] = row.get("amt_financed", "")

    # PF Amount – fixed at ₹5,900
    ws[f"AS{next_row}"] = 5900

    # Interest Amount – 10 % of financed amount
    try:
        financed = float(row.get("amt_financed") or 0)
        interest = round(financed * 0.10)
    except Exception:
        interest = ""
    ws[f"AT{next_row}"] = interest

    # ── Inject validation formulas for rows beyond template pre-fill ─────
    _inject_validation_formulas(ws, next_row)

    # ── Save ─────────────────────────────────────────────────────────────
    wb.save(output_path)

    return next_row


# =========================================================
# READ ALL ROWS FROM EXCEL
# =========================================================

def read_all_rows_from_excel(path: Path) -> List[Dict]:
    """
    Read every data row (row 4+) that has at least one non-blank value
    in the core data columns (A–AT).  Validation columns AU–BC contain
    Excel formula strings when opened without data_only, so we evaluate
    them locally with the simple rules that mirror the sheet formulas.

    Returns a list of dicts keyed by column letter for data columns, plus
    a "validation" sub-dict for AU–BC derived values.
    """
    if not path.exists():
        return []

    # Load WITHOUT data_only so we can read col D values (not formula text)
    # and detect whether validation formula cells are present.
    wb = load_workbook(path, data_only=True)
    ws = wb.active

    # Column letter → friendly key mapping for data columns A–AT
    DATA_COLS = {
        "A": "serial_no", "B": "lot_no", "C": "file_no",
        "D": "rps_no", "E": "receipt_no", "F": "sourcing_branch",
        "G": "first_name", "H": "middle_name", "I": "last_name",
        "J": "dob", "K": "gender", "L": "marital_status",
        "M": "category", "N": "pan_no", "O": "religion",
        "P": "current_residence", "Q": "qualification", "R": "profession",
        "S": "address_1", "T": "address_2", "U": "address_3",
        "V": "city", "W": "state", "X": "zip_code",
        "Y": "phone_1", "Z": "phone_2", "AA": "mobile",
        "AB": "email", "AC": "repayment_mode", "AD": "account_no",
        "AE": "micr", "AF": "ifsc", "AG": "bank_name",
        "AH": "account_type", "AI": "id_proof_no", "AJ": "id_expiry",
        "AK": "id_doc_type", "AL": "address_proof_no",
        "AM": "address_expiry", "AN": "address_doc_type",
        "AO": "sc_st_flag", "AP": "plot_size", "AQ": "asset_cost",
        "AR": "amt_financed", "AS": "pf_amount", "AT": "interest_amt",
    }
    VAL_COLS = {
        "AU": "pan_format", "AV": "pan_duplicate",
        "AW": "mobile_duplicate", "AX": "plot_size_check",
        "AY": "category_check", "AZ": "ifsc_check",
        "BA": "mobile_format", "BB": "rps_no_check",
        "BC": "overall_status",
    }

    rows_out = []

    for r in range(4, ws.max_row + 1):
        # Read data columns
        data = {key: ws[f"{col}{r}"].value for col, key in DATA_COLS.items()}

        # Skip completely blank rows (check key identifier columns)
        if data["rps_no"] in (None, "") and data["receipt_no"] in (None, ""):
            continue

        # Read validation columns (data_only=True returns cached formula result
        # if Excel last saved with calculated values; otherwise None)
        validation = {key: ws[f"{col}{r}"].value for col, key in VAL_COLS.items()}

        # If cached validation values are absent (None), compute them locally
        # so the caller always gets meaningful status information.
        if validation["overall_status"] is None:
            validation = _compute_validation(data)

        rows_out.append({**data, "validation": validation})

    return rows_out


def _compute_validation(data: Dict) -> Dict:
    """
    Pure-Python mirror of the AU–BC Excel formulas.
    Called when the workbook was opened with data_only=True but the
    formula results were not cached (i.e., never opened in real Excel).
    """
    pan = str(data.get("pan_no") or "").strip().upper()
    mobile = data.get("mobile")
    plot = data.get("plot_size")
    category = str(data.get("category") or "").strip().upper()
    ifsc = str(data.get("ifsc") or "").strip()
    repay = str(data.get("repayment_mode") or "").strip().upper()
    rps = str(data.get("rps_no") or "").strip()

    # AU – PAN format (5 alpha + P + 1 alpha + 4 digits + 1 alpha)
    pan_ok = bool(re.match(r'^[A-Z]{3}P[A-Z]\d{4}[A-Z]$', pan)) if pan else None
    au = "✔ Valid PAN" if pan_ok else ("✘ Invalid PAN" if pan else "")

    # AV – PAN duplicate: can't check without full dataset, mark unknown
    av = "⚠ Dup check N/A"

    # AW – Mobile duplicate: same limitation
    aw = "⚠ Dup check N/A"

    # AX – Plot size
    valid_sizes = {162, 183, 184, 200, 223, 290}
    try:
        ax = "✔ Valid Size" if int(float(plot)) in valid_sizes else "✘ Invalid Size"
    except Exception:
        ax = "" if not plot else "✘ Invalid Size"

    # AY – Category
    ay = "✔ Valid Cat" if category in {"GEN", "SC/ST", "SC", "ST"} else ("✘ Invalid Cat" if category else "")

    # AZ – IFSC / repayment mode
    if not repay and not ifsc:
        az = ""
    elif repay == "AUTO DEBIT":
        az = "✔ Valid IFSC"
    else:
        az = "✔ Valid IFSC" if len(ifsc) == 11 else "✘ Invalid IFSC"

    # BA – Mobile format (10 digits)
    try:
        mob_str = str(int(float(mobile))) if mobile is not None else ""
        ba = "✔ Valid Mob" if len(mob_str) == 10 else "✘ Invalid Mob"
    except Exception:
        ba = "" if not mobile else "✘ Invalid Mob"

    # BB – RPS No (12 chars, uppercase alphanumeric)
    bb = "✔ Valid RPS" if (len(rps) == 12 and rps == rps.upper()) else ("✘ RPS Must Be 12 Chars" if rps else "")

    # BC – Overall status
    errors = [v for v in [au, av, aw, ax, ay, az, ba, bb]
              if any(kw in str(v) for kw in ("DUPLICATE", "ERROR", "✗", "Invalid", "✘"))]
    bc = "✓ READY TO SUBMIT" if (pan or plot) and not errors else ("✗ ERRORS – CHECK" if errors else "")

    return {
        "pan_format": au, "pan_duplicate": av, "mobile_duplicate": aw,
        "plot_size_check": ax, "category_check": ay, "ifsc_check": az,
        "mobile_format": ba, "rps_no_check": bb, "overall_status": bc,
    }


# =========================================================
# XML EXPORT
# =========================================================

def write_xml_from_excel(xlsx_path: Path, xml_path: Path) -> None:
    rows = read_all_rows_from_excel(xlsx_path)

    root = ET.Element("RPSApplications")
    root.set("generated_at", datetime.now().isoformat())
    root.set("count", str(len(rows)))

    for idx, row_data in enumerate(rows, start=1):
        app = ET.SubElement(root, "Application")
        app.set("row", str(idx))

        # Flatten validation sub-dict into the element
        validation = row_data.pop("validation", {})
        for key, value in row_data.items():
            child = ET.SubElement(app, key)
            child.text = "" if value is None else str(value)

        val_el = ET.SubElement(app, "validation")
        for key, value in validation.items():
            child = ET.SubElement(val_el, key)
            child.text = "" if value is None else str(value)

    ensure_parent(xml_path)
    tree = ET.ElementTree(root)
    ET.indent(tree, space="  ")
    tree.write(xml_path, encoding="utf-8", xml_declaration=True)


# =========================================================
# JSON LOG
# =========================================================

def append_json_log(row: Dict, json_path: Path) -> None:
    ensure_parent(json_path)

    existing = []
    if json_path.exists():
        try:
            existing = json.loads(json_path.read_text(encoding="utf-8"))
        except Exception:
            existing = []

    existing.append({"saved_at": datetime.now().isoformat(), **row})
    json_path.write_text(
        json.dumps(existing, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )