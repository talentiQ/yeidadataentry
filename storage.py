from __future__ import annotations

import json
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path
from shutil import copyfile
from typing import Dict, List

from openpyxl import load_workbook


# =========================================================
# TEMPLATE FILE
# =========================================================

BASE_DIR = Path(__file__).resolve().parent

TEMPLATE_FILE = BASE_DIR / "data" / "RPS_10_MASTER.xlsx"


# =========================================================
# ENSURE DIRECTORY
# =========================================================

def ensure_parent(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)


# =========================================================
# APPEND TO EXCEL
# =========================================================

def append_to_excel(row: Dict, output_path: Path) -> int:

    ensure_parent(output_path)

    # Create output from template first time
    if not output_path.exists():

        copyfile(
            TEMPLATE_FILE,
            output_path,
        )

    wb = load_workbook(output_path)

    ws = wb.active

    # =====================================================
    # FIND FIRST EMPTY ROW
    # =====================================================

    next_row = 4

    while ws[f"E{next_row}"].value not in (None, ""):
        next_row += 1

    # =====================================================
    # SERIAL NUMBER
    # =====================================================

    ws[f"A{next_row}"] = next_row - 3

    # =====================================================
    # BASIC DETAILS
    # =====================================================

    ws[f"B{next_row}"] = row.get("lot_no", "")
    ws[f"C{next_row}"] = row.get("file_no", "")
    ws[f"D{next_row}"] = row.get("rps_no", "")
    ws[f"E{next_row}"] = row.get("receipt_no", "")
    ws[f"F{next_row}"] = row.get("sourcing_branch", "")

    # =====================================================
    # NAME
    # =====================================================

    ws[f"G{next_row}"] = row.get("first_name", "")
    ws[f"H{next_row}"] = row.get("middle_name", "")
    ws[f"I{next_row}"] = row.get("last_name", "")

    # =====================================================
    # PERSONAL DETAILS
    # =====================================================

    ws[f"J{next_row}"] = row.get("dob", "")
    ws[f"K{next_row}"] = row.get("gender", "")
    ws[f"L{next_row}"] = row.get("marital_status", "")
    ws[f"M{next_row}"] = row.get("category", "")
    ws[f"N{next_row}"] = row.get("pan_no", "")

    ws[f"O{next_row}"] = row.get("religion", "")
    ws[f"P{next_row}"] = row.get("current_residence", "")
    ws[f"Q{next_row}"] = row.get("qualification", "")
    ws[f"R{next_row}"] = row.get("profession", "")

    # =====================================================
    # ADDRESS
    # =====================================================

    ws[f"S{next_row}"] = row.get("address_1", "")
    ws[f"T{next_row}"] = row.get("address_2", "")
    ws[f"U{next_row}"] = row.get("address_3", "")

    ws[f"V{next_row}"] = row.get("city", "")
    ws[f"W{next_row}"] = row.get("state", "")
    ws[f"X{next_row}"] = row.get("zip_code", "")

    # =====================================================
    # CONTACT
    # =====================================================

    ws[f"Y{next_row}"] = row.get("phone_1", "")
    ws[f"Z{next_row}"] = row.get("phone_2", "")
    ws[f"AA{next_row}"] = row.get("mobile", "")

    ws[f"AB{next_row}"] = row.get("email", "")

    # =====================================================
    # BANK DETAILS
    # =====================================================

    ws[f"AC{next_row}"] = row.get("repayment_mode", "")

    ws[f"AD{next_row}"] = row.get("account_no", "")
    ws[f"AE{next_row}"] = row.get("micr", "")
    ws[f"AF{next_row}"] = row.get("ifsc", "")
    ws[f"AG{next_row}"] = row.get("bank_name", "")

    ws[f"AH{next_row}"] = row.get("account_type", "")

    # =====================================================
    # KYC
    # =====================================================

    ws[f"AI{next_row}"] = row.get("id_proof_no", "")
    ws[f"AJ{next_row}"] = row.get("id_expiry", "")
    ws[f"AK{next_row}"] = row.get("id_doc_type", "")

    ws[f"AL{next_row}"] = row.get("address_proof_no", "")
    ws[f"AM{next_row}"] = row.get("address_expiry", "")
    ws[f"AN{next_row}"] = row.get("address_doc_type", "")

    # =====================================================
    # SC/ST
    # =====================================================

    ws[f"AO{next_row}"] = row.get("sc_st_flag", "")

    # =====================================================
    # PROPERTY
    # =====================================================

    ws[f"AP{next_row}"] = row.get("plot_size", "")

    ws[f"AQ{next_row}"] = row.get("asset_cost", "")

    ws[f"AR{next_row}"] = row.get("amt_financed", "")

    # PF Amount Fixed
    ws[f"AS{next_row}"] = 5900

    # Interest Amount
    try:

        financed = float(
            row.get("amt_financed") or 0
        )

        interest = round(financed * 0.10)

    except Exception:

        interest = ""

    ws[f"AT{next_row}"] = interest

    # =====================================================
    # SAVE
    # =====================================================

    wb.save(output_path)

    return next_row


# =========================================================
# READ ALL ROWS
# =========================================================

def read_all_rows_from_excel(path: Path) -> List[Dict]:

    if not path.exists():
        return []

    wb = load_workbook(
        path,
        data_only=True,
    )

    ws = wb.active

    rows = []

    for row in ws.iter_rows(
        min_row=4,
        values_only=True,
    ):

        if not any(row):
            continue

        rows.append(list(row))

    return rows


# =========================================================
# XML EXPORT
# =========================================================

def write_xml_from_excel(
    xlsx_path: Path,
    xml_path: Path,
) -> None:

    rows = read_all_rows_from_excel(
        xlsx_path
    )

    root = ET.Element("RPSApplications")

    root.set(
        "generated_at",
        datetime.now().isoformat(),
    )

    root.set(
        "count",
        str(len(rows)),
    )

    for idx, row in enumerate(rows, start=1):

        app = ET.SubElement(
            root,
            "Application",
        )

        app.set("row", str(idx))

        for col_idx, value in enumerate(row, start=1):

            tag = f"col_{col_idx}"

            child = ET.SubElement(
                app,
                tag,
            )

            child.text = (
                ""
                if value is None
                else str(value)
            )

    ensure_parent(xml_path)

    tree = ET.ElementTree(root)

    ET.indent(tree, space="  ")

    tree.write(
        xml_path,
        encoding="utf-8",
        xml_declaration=True,
    )


# =========================================================
# JSON LOG
# =========================================================

def append_json_log(
    row: Dict,
    json_path: Path,
) -> None:

    ensure_parent(json_path)

    existing = []

    if json_path.exists():

        try:

            existing = json.loads(
                json_path.read_text(
                    encoding="utf-8"
                )
            )

        except Exception:

            existing = []

    existing.append({
        "saved_at": datetime.now().isoformat(),
        **row,
    })

    json_path.write_text(
        json.dumps(
            existing,
            indent=2,
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )