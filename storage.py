from __future__ import annotations

import json
import os
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path
from typing import Dict, List

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


HEADERS = [
    "Entry DateTime",
    "Receipt No",
    "SOL ID",
    "Account Type",
    "Employment Type",
    "Form No",
    "Application No",
    "Applicant Name",
    "Father/Husband Name",
    "Category",
    "Sub Category",
    "Plot Size",
    "Gender",
    "Age",
    "Date of Birth",
    "Marital Status",
    "Payment Plan",
    "Registration Money",
    "Deposit Amount",
    "Asset Cost",
    "Amt Financed",
    "Customer Contribution",
    "Form Fees",
    "Paid By",
    "Mobile No",
    "Email",
    "Mailing Address 1",
    "Mailing Address 2",
    "Mailing Address 3",
    "Mailing City",
    "Mailing Pincode",
    "Permanent Address 1",
    "Permanent Address 2",
    "Permanent Address 3",
    "Permanent City",
    "Permanent Pincode",
    "Account Name",
    "Bank Name",
    "Branch Name",
    "Account Number",
    "IFSC Code",
    "PAN Number",
    "ID Proof Last Digit",
    "Printed Receipt No",
    "Handwritten Receipt No",
    "Payment Date",
    "DD/UTR No",
    "Payment ID",
    "EMD Branch",
    "Co-Applicant Name",
    "Relation With Applicant",
    "Co-Father/Husband Name",
    "Handicapped",
    "Notes",
]

FIELD_MAP = {
    "Receipt No": "receipt_no",
    "SOL ID": "sol_id",
    "Account Type": "account_type",
    "Employment Type": "employment_type",
    "Form No": "form_no",
    "Application No": "application_no",
    "Applicant Name": "applicant_name",
    "Father/Husband Name": "father_or_husband_name",
    "Category": "category",
    "Sub Category": "sub_category",
    "Plot Size": "plot_size",
    "Gender": "gender",
    "Age": "age",
    "Date of Birth": "date_of_birth",
    "Marital Status": "marital_status",
    "Payment Plan": "payment_plan",
    "Registration Money": "registration_money",
    "Deposit Amount": "deposit_amount",
    "Asset Cost": "asset_cost",
    "Amt Financed": "amt_financed",
    "Customer Contribution": "customer_contribution",
    "Form Fees": "form_fees",
    "Paid By": "paid_by",
    "Mobile No": "mobile_no",
    "Email": "email",
    "Mailing Address 1": "mailing_address_1",
    "Mailing Address 2": "mailing_address_2",
    "Mailing Address 3": "mailing_address_3",
    "Mailing City": "mailing_city",
    "Mailing Pincode": "mailing_pincode",
    "Permanent Address 1": "permanent_address_1",
    "Permanent Address 2": "permanent_address_2",
    "Permanent Address 3": "permanent_address_3",
    "Permanent City": "permanent_city",
    "Permanent Pincode": "permanent_pincode",
    "Account Name": "account_name",
    "Bank Name": "bank_name",
    "Branch Name": "branch_name",
    "Account Number": "account_number",
    "IFSC Code": "ifsc_code",
    "PAN Number": "pan_number",
    "ID Proof Last Digit": "id_proof_no",
    "Printed Receipt No": "printed_receipt_no",
    "Handwritten Receipt No": "handwritten_receipt_no",
    "Payment Date": "payment_date",
    "DD/UTR No": "dd_utr_no",
    "Payment ID": "payment_id",
    "EMD Branch": "emd_branch",
    "Co-Applicant Name": "co_applicant_name",
    "Relation With Applicant": "relation_with_applicant",
    "Co-Father/Husband Name": "co_father_or_husband_name",
    "Handicapped": "handicapped",
    "Notes": "notes",
}

NUMERIC_HEADERS = {
    "Registration Money",
    "Deposit Amount",
    "Asset Cost",
    "Amt Financed",
    "Customer Contribution",
    "Form Fees",
}


def ensure_parent(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)


def create_workbook(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "YEIDA Entries"
    ws.append(HEADERS)
    style_sheet(ws)
    ensure_parent(path)
    wb.save(path)


def style_sheet(ws) -> None:
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(color="FFFFFF", bold=True)
    thin = Side(style="thin", color="D9E2F3")
    border = Border(top=thin, left=thin, right=thin, bottom=thin)

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(HEADERS))}1"

    widths = {
        "A": 20, "B": 12, "C": 10, "D": 12, "E": 16,
        "F": 16, "G": 16, "H": 24, "I": 24, "J": 18,
        "K": 18, "L": 10, "M": 10, "N": 8, "O": 14,
        "P": 14, "Q": 14, "R": 16, "S": 14, "T": 14,
        "U": 14, "V": 18, "W": 12, "X": 12, "Y": 14,
        "Z": 26,
    }
    for col_idx in range(1, len(HEADERS) + 1):
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = widths.get(col_letter, 18)


def append_to_excel(row: Dict, path: Path) -> int:
    if not path.exists():
        create_workbook(path)

    wb = load_workbook(path)
    ws = wb.active

    # Repair headers if workbook was manually changed
    current_headers = [ws.cell(row=1, column=i).value for i in range(1, len(HEADERS) + 1)]
    if current_headers != HEADERS:
        for i, header in enumerate(HEADERS, 1):
            ws.cell(row=1, column=i, value=header)
        style_sheet(ws)

    entry = []
    for header in HEADERS:
        if header == "Entry DateTime":
            entry.append(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        else:
            entry.append(row.get(FIELD_MAP.get(header, ""), ""))

    ws.append(entry)
    row_num = ws.max_row

    # Formatting new row
    thin = Side(style="thin", color="D9E2F3")
    border = Border(top=thin, left=thin, right=thin, bottom=thin)
    for col_idx, header in enumerate(HEADERS, 1):
        cell = ws.cell(row=row_num, column=col_idx)
        cell.border = border
        cell.alignment = Alignment(vertical="top", wrap_text=True)
        if header in NUMERIC_HEADERS:
            cell.number_format = '#,##0.00'

    style_sheet(ws)
    ensure_parent(path)
    wb.save(path)
    return row_num


def read_all_rows_from_excel(path: Path) -> List[Dict]:
    if not path.exists():
        return []
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(row):
            continue
        rows.append({HEADERS[i]: row[i] for i in range(min(len(HEADERS), len(row)))})
    return rows


def write_xml_from_excel(xlsx_path: Path, xml_path: Path) -> None:
    rows = read_all_rows_from_excel(xlsx_path)
    root = ET.Element("YEIDAApplications")
    root.set("scheme", "RPS-10 2026")
    root.set("generated_at", datetime.now().isoformat(timespec="seconds"))
    root.set("count", str(len(rows)))

    for idx, row in enumerate(rows, start=1):
        app = ET.SubElement(root, "Application")
        app.set("row", str(idx))
        for header in HEADERS:
            tag = header.lower().replace("/", "_").replace(" ", "_").replace("-", "_")
            tag = ''.join(ch for ch in tag if ch.isalnum() or ch == '_')
            child = ET.SubElement(app, tag)
            value = row.get(header)
            child.text = "" if value is None else str(value)

    ensure_parent(xml_path)
    tree = ET.ElementTree(root)
    ET.indent(tree, space="  ")
    tree.write(xml_path, encoding="utf-8", xml_declaration=True)


def append_json_log(row: Dict, json_path: Path) -> None:
    ensure_parent(json_path)
    existing = []
    if json_path.exists():
        try:
            existing = json.loads(json_path.read_text(encoding="utf-8"))
        except Exception:
            existing = []
    existing.append({"saved_at": datetime.now().isoformat(timespec="seconds"), **row})
    json_path.write_text(json.dumps(existing, indent=2, ensure_ascii=False), encoding="utf-8")
