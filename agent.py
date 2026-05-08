from __future__ import annotations

import base64
import mimetypes
import os
import re
from pathlib import Path
from typing import List, Optional

from openai import OpenAI
from pydantic import BaseModel, Field


# =========================================================
# OCR MODEL
# =========================================================

class YEIDAExtract(BaseModel):

    form_no: Optional[str] = None
    application_no: Optional[str] = None

    applicant_name: Optional[str] = None
    father_or_husband_name: Optional[str] = None

    category: Optional[str] = None
    sub_category: Optional[str] = None

    plot_size: Optional[int] = None

    date_of_birth: Optional[str] = None

    gender: Optional[str] = None
    marital_status: Optional[str] = None

    mobile_no: Optional[str] = None
    email: Optional[str] = None

    mailing_address_1: Optional[str] = None
    mailing_address_2: Optional[str] = None
    mailing_address_3: Optional[str] = None

    mailing_city: Optional[str] = None
    mailing_pincode: Optional[str] = None

    account_name: Optional[str] = None
    bank_name: Optional[str] = None
    branch_name: Optional[str] = None
    account_number: Optional[str] = None
    ifsc_code: Optional[str] = None

    pan_number: Optional[str] = None

    receipt_no: Optional[str] = None

    handwritten_receipt_no: Optional[str] = None
    printed_receipt_no: Optional[str] = None

    sol_id: Optional[str] = None

    account_type: Optional[str] = None

    employment_type: Optional[str] = None

    amt_financed: Optional[float] = None
    asset_cost: Optional[float] = None

    notes: Optional[str] = None


# =========================================================
# SYSTEM PROMPT
# =========================================================

SYSTEM_PROMPT = """
You are an OCR extraction agent for YEIDA RPS-10/2026 housing scheme documents.

You will receive multiple images from ONE applicant's document set. These may include:
1. YEIDA Application Form (APPLICATION-FORM-CUM-FACILITY AGREEMENT) — main multi-page form
2. ICICI Bank Documents Receipt (pink slip) — has the ICICI booking receipt number
3. YEIDA Online Payment Receipt — shows Application No, mobile, name, EMD/UPF amounts
4. Aadhaar card (front: name, DOB, gender; back: address with W/O or S/O guardian name)
5. PAN card — has PAN number, name, father/husband name, DOB
6. Cancelled cheque — has account number, IFSC, bank name, account holder name
7. NACH / Security Mandate form — has account number, IFSC, bank name

Extract each field from the best available source:

APPLICATION DETAILS:
- application_no: 12-character YEIDA application/form number (format RPS10XXXXXXXX).
  Find on YEIDA payment receipt as "Application No" or "Form No", or on the application
  form header as "YEIDA From No." — e.g., RPS100225128.
- receipt_no: ICICI Bank booking receipt number — short numeric code (3-6 digits) printed
  at the bottom-right of the pink ICICI Documents Receipt slip. NOT the YEIDA receipt number.
- handwritten_receipt_no: Same receipt number if it appears handwritten on the ICICI slip.
- printed_receipt_no: Same receipt number if it appears printed on the ICICI slip.
- sol_id: Branch SOL ID if visible on the ICICI receipt slip.

APPLICANT PERSONAL DETAILS (from Application Form or Aadhaar):
- applicant_name: Full name in CAPITALS exactly as written (e.g., "SHIKHA GOYAL").
- father_or_husband_name: Father's name or spouse name as written on the form.
- date_of_birth: Format dd/mm/yyyy — read from Aadhaar front or PAN card.
- gender: MALE or FEMALE
- marital_status: SINGLE, MARRIED, or OTHER
- category: Exactly as ticked on the form — "GENERAL AND OTHERS", "SC", "ST", "OBC", or "EWS".
- pan_number: Exactly 10 uppercase characters from the PAN card image (e.g., ALKPG2479G).
  NEVER guess or invent a PAN number.
- employment_type: As ticked on the form — SALARIED, PROFESSIONAL, or SELF EMPLOYED.
- religion: As ticked — Hindu, Muslim, Sikh, Christian, Buddhist, or Others.

PLOT SIZE:
- plot_size: Plot size as an integer (square meters). Look for a plot size field on the
  application form first. If not visible, infer from the YEIDA EMD receipt total amount:
  ₹52,361 → 162 sqm | ₹59,723 → 183 sqm | ₹60,047 → 184 sqm |
  ₹72,520 → 200 sqm | ₹72,774 → 223 sqm | ₹94,639 → 290 sqm

ADDRESS (from Application Form or Aadhaar back):
- mailing_address_1: House/flat number and street name.
- mailing_address_2: Area / colony / sector (if present).
- mailing_address_3: Landmark (if present).
- mailing_city: City name.
- mailing_pincode: 6-digit PIN code.

CONTACT:
- mobile_no: 10-digit mobile from the application form or YEIDA payment receipt.
  Return digits ONLY — no spaces, dashes, or country codes.
- email: Email address exactly as written on the application form.

BANK DETAILS (from cancelled cheque or NACH mandate — prefer cheque):
- account_number: Full account number as text, preserving any leading zeros
  (e.g., "025401513496"). NEVER guess.
- ifsc_code: Exactly 11-character IFSC from the cheque (e.g., ICIC0000254). NEVER guess.
- bank_name: Bank name (e.g., "ICICI BANK").
- account_type: SAVINGS or CURRENT.

FINANCIAL:
- amt_financed: The ICICI facility amount in rupees from the NACH mandate or application form
  (e.g., 72520.0). This is the EMD loan amount ICICI will disburse.
- asset_cost: Leave null — auto-computed from plot size by the system.

Strict rules:
- NEVER fabricate or guess PAN numbers, account numbers, IFSC codes, or mobile numbers.
- Aadhaar numbers are always masked in scanned copies — do NOT attempt to read them.
- Return null for any field that is not clearly visible in the images.
- mobile_no: digits only, no dashes or spaces.
- pan_number: uppercase only.
- For receipt_no: use the short number from the ICICI pink booking slip, NOT the long YEIDA
  online payment receipt numbers (which are 6 digits like 457507).
"""


# =========================================================
# IMAGE ENCODING
# =========================================================

def encode_image_to_data_url(path: Path) -> str:

    mime, _ = mimetypes.guess_type(path)

    if not mime:
        mime = "image/jpeg"

    encoded = base64.b64encode(
        path.read_bytes()
    ).decode("utf-8")

    return f"data:{mime};base64,{encoded}"


# =========================================================
# OCR EXTRACTION
# =========================================================

def extract_from_images(
    image_paths: List[Path],
    model: Optional[str] = None,
) -> YEIDAExtract:

    if not image_paths:
        raise ValueError("No images uploaded")

    client = OpenAI(
        api_key=os.environ.get("OPENAI_API_KEY"),
        timeout=60.0,
    )

    model = model or os.environ.get(
        "OPENAI_MODEL",
        "gpt-4.1-mini",
    )

    content = [
        {
            "type": "input_text",
            "text": "Extract applicant data from these YEIDA RPS-10/2026 scheme document images.",
        }
    ]

    for path in image_paths:

        content.append({
            "type": "input_image",
            "image_url": encode_image_to_data_url(path),
        })

    response = client.responses.parse(
        model=model,
        input=[
            {
                "role": "system",
                "content": SYSTEM_PROMPT,
            },
            {
                "role": "user",
                "content": content,
            },
        ],
        text_format=YEIDAExtract,
    )

    parsed = response.output_parsed

    if not parsed:
        raise RuntimeError("OCR extraction failed — no structured output returned")

    return parsed


# =========================================================
# HELPERS
# =========================================================

def only_digits(value):

    if not value:
        return None

    digits = re.sub(r"\D", "", str(value))

    return digits or None


def normalize_text(value):

    if not value:
        return None

    value = re.sub(r"\s+", " ", str(value))

    return value.strip()


def normalize_category(value):
    """
    Normalise raw OCR category text to the values the Excel AY formula
    accepts: GEN · SC · ST · SC/ST.

    Excel AY formula checks:
      UPPER(TRIM(M)) = "GEN"  |  "SC/ST"  |  "SC"  |  "ST"

    OBC and EWS are not in the Excel AY valid set — mapped to GEN.
    The original raw category is preserved in the applicant notes via
    apply_business_rules so no information is silently lost.
    """
    text = normalize_text(value)

    if not text:
        return None

    upper = text.upper()

    # Explicit combined check first to avoid SC/ST being split
    if "SC/ST" in upper:
        return "SC/ST"

    if "SC" in upper:
        return "SC"

    if "ST" in upper:
        return "ST"

    # OBC and EWS → GEN (only valid general-category value for AY formula)
    return "GEN"


def normalize_account_type(value, override=None):
    """
    Convert any account-type text (OCR or override) to the numeric code
    that column AH in the YEIDA master sheet expects.

    FIX: The improved OCR now correctly reads "SAVINGS" / "CURRENT" from
    the NACH mandate form, but the Excel column AH uses numeric codes:
      31 = Savings Bank
      32 = Current Account
    Without this normalisation, "SAVINGS" was being written to AH verbatim.
    """
    raw = str(override or value or "").strip().upper()

    SAVINGS_ALIASES = {"31", "SAVINGS", "SB", "SAVING", "SAVINGS BANK", "SBACCOUNT"}
    CURRENT_ALIASES = {"32", "10", "CURRENT", "CA", "CURRENT ACCOUNT"}

    if raw in SAVINGS_ALIASES:
        return "31"
    if raw in CURRENT_ALIASES:
        return "32"
    return "31"   # default to Savings when unknown


# =========================================================
# NAME SPLIT
# =========================================================

def split_name(full_name):

    full_name = normalize_text(full_name)

    if not full_name:
        return "", "", ""

    parts = full_name.split()

    if len(parts) == 1:
        return parts[0], "", ""

    if len(parts) == 2:
        return parts[0], "", parts[1]

    return (
        parts[0],
        " ".join(parts[1:-1]),
        parts[-1],
    )


# =========================================================
# PRICING
# =========================================================
# Source: YEIDA Rate Master sheet — ICICI Bank YEIDA RPS-10/2026
# Columns: plot_size → earnest_money, bank_loan, pf, interest_3m
# Interest = Bank Loan × 11% p.a. × 3 months = Bank Loan × 0.0275
# DO NOT edit these values manually — update from the Rate Master sheet.

PRICING = {
    "GEN": {
        162: {"earnest_money": 587412,  "amt_financed": 528671, "pf": 5900, "interest": 14538},
        183: {"earnest_money": 663588,  "amt_financed": 597229, "pf": 5900, "interest": 16424},
        184: {"earnest_money": 667184,  "amt_financed": 600466, "pf": 5900, "interest": 16513},
        200: {"earnest_money": 725200,  "amt_financed": 652680, "pf": 5900, "interest": 17949},
        223: {"earnest_money": 808598,  "amt_financed": 727738, "pf": 5900, "interest": 20013},
        290: {"earnest_money": 1051540, "amt_financed": 946386, "pf": 5900, "interest": 26026},
    },
    "SC/ST": {
        162: {"earnest_money": 293706,  "amt_financed": 264335, "pf": 5900, "interest": 7269},
        183: {"earnest_money": 331779,  "amt_financed": 298601, "pf": 5900, "interest": 8212},
        184: {"earnest_money": 333592,  "amt_financed": 300233, "pf": 5900, "interest": 8256},
        200: {"earnest_money": 362600,  "amt_financed": 326340, "pf": 5900, "interest": 8974},
        223: {"earnest_money": 404299,  "amt_financed": 363869, "pf": 5900, "interest": 10006},
        290: {"earnest_money": 525770,  "amt_financed": 473193, "pf": 5900, "interest": 13013},
    },
}

def _get_pricing(plot_size: int, category: str) -> dict:
    """
    Look up rate master values for a given plot size and resolved category.
    SC, ST, SC/ST all use the SC/ST table. Everything else uses GEN.
    Returns an empty dict if plot_size is not a valid YEIDA size.
    """
    table_key = "SC/ST" if category in ("SC", "ST", "SC/ST") else "GEN"
    return PRICING.get(table_key, {}).get(plot_size, {})


# =========================================================
# BUSINESS RULES
# =========================================================

def apply_business_rules(
    data: YEIDAExtract,
    account_type_override=None,
    category_override=None,
    receipt_override=None,
    sol_id_override=None,
    plot_size_override=None,       # FIX: was missing — manual override for plot size
):

    extracted = data.model_dump()

    # Normalize all string fields
    for k, v in extracted.items():
        if isinstance(v, str):
            extracted[k] = normalize_text(v)

    # Name split
    first_name, middle_name, last_name = split_name(
        extracted.get("applicant_name")
    )

    # Plot size — FIX: apply override before PRICING lookup
    raw_plot = plot_size_override or extracted.get("plot_size")
    try:
        plot_size = int(raw_plot)
    except Exception:
        plot_size = None

    # Resolved category — must come BEFORE pricing lookup
    raw_category = category_override or extracted.get("category")
    resolved_category = normalize_category(raw_category)

    # SC/ST flag — 'Y'/'N' to match template column AO
    sc_st_flag = "Y" if resolved_category in ("SC", "ST", "SC/ST") else "N"

    # Pricing — pulled directly from YEIDA Rate Master for the correct category table.
    # SC/ST rates are pre-computed in the rate master (NOT a simple 50% of GEN —
    # they differ due to YEIDA scheme rules). GEN rates apply for all other categories.
    pricing = _get_pricing(plot_size, resolved_category) if plot_size else {}
    asset_cost   = pricing.get("earnest_money")   # Earnest Money from rate master
    amt_financed = pricing.get("amt_financed")    # Bank Loan from rate master
    interest_amt = pricing.get("interest")        # Interest 3M @11% from rate master

    # FIX: mobile must be stored as int so Excel ISNUMBER(AA{R}) = TRUE.
    raw_mobile = only_digits(extracted.get("mobile_no"))
    try:
        mobile = int(raw_mobile) if raw_mobile else None
    except (ValueError, TypeError):
        mobile = None

    # Build notes — FIX: sol_id_override is now recorded instead of silently dropped.
    # Also flag if OCR category was remapped (OBC/EWS → GEN) so data entry can verify.
    notes_parts = []
    if sol_id_override:
        notes_parts.append(f"SOL_ID: {sol_id_override}")
    if raw_category and resolved_category == "GEN" and raw_category.upper() not in ("GEN", "GENERAL", "GENERAL AND OTHERS"):
        notes_parts.append(f"OCR_CATEGORY: {raw_category} (mapped to GEN)")
    if extracted.get("notes"):
        notes_parts.append(extracted["notes"])
    notes_value = " | ".join(notes_parts) or None

    row = {

        # Basic
        "lot_no": "",
        "file_no": "",

        "rps_no": (
            extracted.get("application_no")
            or extracted.get("form_no")
        ),

        # Receipt: prefer override > handwritten > generic > printed
        # Kept as string (only_digits preserves leading zeros correctly)
        "receipt_no": (
            only_digits(receipt_override)
            or only_digits(extracted.get("handwritten_receipt_no"))
            or only_digits(extracted.get("receipt_no"))
            or only_digits(extracted.get("printed_receipt_no"))
        ),

        "sourcing_branch": "",

        # Name
        "first_name": first_name,
        "middle_name": middle_name,
        "last_name": last_name,

        # Personal
        "dob": extracted.get("date_of_birth"),
        "gender": extracted.get("gender"),
        "marital_status": extracted.get("marital_status"),
        "category": resolved_category,
        "pan_no": extracted.get("pan_number"),
        "religion": "",
        "current_residence": "",
        "qualification": "",
        "profession": extracted.get("employment_type") or "SALARIED",

        # Address
        "address_1": extracted.get("mailing_address_1"),
        "address_2": extracted.get("mailing_address_2"),
        "address_3": extracted.get("mailing_address_3"),
        "city": extracted.get("mailing_city"),
        "state": "",
        "zip_code": extracted.get("mailing_pincode"),

        # Contact
        "phone_1": "",
        "phone_2": "",
        "mobile": mobile,   # int so Excel ISNUMBER() = TRUE
        "email": extracted.get("email"),

        # Bank
        "repayment_mode": "AUTO DEBIT",
        "account_no": extracted.get("account_number"),
        "micr": "",
        "ifsc": extracted.get("ifsc_code"),
        "bank_name": extracted.get("bank_name"),
        # FIX: normalize to numeric code ("31"/"32") — OCR may return "SAVINGS"/"CURRENT"
        "account_type": normalize_account_type(
            extracted.get("account_type"),
            override=account_type_override,
        ),

        # KYC
        "id_proof_no": "",
        "id_expiry": "",
        "id_doc_type": "",
        "address_proof_no": "",
        "address_expiry": "",
        "address_doc_type": "",

        # SC/ST Flag — 'Y'/'N' to match template column AO
        "sc_st_flag": sc_st_flag,

        # Property — values from YEIDA Rate Master (category + plot_size specific)
        "plot_size": plot_size,
        "asset_cost": asset_cost,        # Earnest Money from rate master
        "amt_financed": amt_financed,    # Bank Loan (90% of EMD) from rate master
        "interest_amt": interest_amt,    # 3-month interest @11% p.a. from rate master

        # Notes — FIX: sol_id and category remapping now recorded here
        "notes": notes_value,
    }

    return row