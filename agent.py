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
You are an OCR extraction system for YEIDA RPS forms.

Extract clean structured data from uploaded form images.

Rules:
- Never guess PAN/mobile/account numbers.
- Prefer handwritten receipt number.
- Mobile should contain digits only.
- PAN should be uppercase.
- Category values:
  GENERAL AND OTHERS
  SC
  ST
  OBC
  EWS
- Return null if not visible.
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

    # FIX: added timeout=60.0 to prevent hanging on large images
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
            "text": "Extract applicant data from YEIDA RPS form images.",
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
        raise RuntimeError("OCR extraction failed")

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

    FIX: was returning 'GENERAL AND OTHERS' which always failed AY check.
    FIX: added explicit 'SC/ST' check before individual SC/ST checks to
         prevent 'SC/ST' input being mis-classified as just 'SC'.
    """
    text = normalize_text(value)

    if not text:
        return None

    upper = text.upper()

    # Explicit combined check first
    if "SC/ST" in upper:
        return "SC/ST"

    if "SC" in upper:
        return "SC"

    if "ST" in upper:
        return "ST"

    # OBC and EWS are not in the Excel AY valid set — map to GEN
    # FIX: was returning 'GENERAL AND OTHERS'; must be 'GEN' for AY formula
    return "GEN"


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

PRICING = {
    162: {
        "asset_cost": 5874120,
        "amt_financed": 528671,
    },
    183: {
        "asset_cost": 6635880,
        "amt_financed": 597229,
    },
    184: {
        "asset_cost": 6671840,
        "amt_financed": 600466,
    },
    200: {
        "asset_cost": 7252000,
        "amt_financed": 652680,
    },
    223: {
        "asset_cost": 8085980,
        "amt_financed": 727738,
    },
    290: {
        "asset_cost": 10515400,
        "amt_financed": 946386,
    },
}


# =========================================================
# BUSINESS RULES
# =========================================================

def apply_business_rules(
    data: YEIDAExtract,
    account_type_override=None,
    category_override=None,
    receipt_override=None,
    sol_id_override=None,
):

    extracted = data.model_dump()

    # Normalize strings
    for k, v in extracted.items():

        if isinstance(v, str):
            extracted[k] = normalize_text(v)

    # Name split
    first_name, middle_name, last_name = split_name(
        extracted.get("applicant_name")
    )

    # Plot size
    plot_size = extracted.get("plot_size")

    try:
        plot_size = int(plot_size)
    except Exception:
        plot_size = None

    # Pricing
    asset_cost = None
    amt_financed = None

    if plot_size in PRICING:

        asset_cost = PRICING[plot_size]["asset_cost"]
        amt_financed = PRICING[plot_size]["amt_financed"]

    # Resolved category
    resolved_category = normalize_category(
        category_override or extracted.get("category")
    )

    # FIX: sc_st_flag was 'YES'/'NO' but template column AO uses 'Y'/'N'
    sc_st_flag = (
        "Y" if resolved_category in ("SC", "ST", "SC/ST") else "N"
    )

    # FIX: mobile must be stored as int so Excel ISNUMBER(AA{R}) = TRUE.
    # only_digits() returns a string; wrapping in int() makes openpyxl
    # write a numeric cell, which Excel recognises as a number.
    raw_mobile = only_digits(extracted.get("mobile_no"))
    try:
        mobile = int(raw_mobile) if raw_mobile else None
    except (ValueError, TypeError):
        mobile = None

    row = {

        # Basic
        "lot_no": "",
        "file_no": "",

        "rps_no": (
            extracted.get("application_no")
            or extracted.get("form_no")
        ),

        # Receipt: prefer override > handwritten > generic > printed
        # only_digits() preserves leading zeros as a string, which is correct
        # (storage writes it as-is; never cast to int)
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
        "mobile": mobile,   # FIX: now int, not string
        "email": extracted.get("email"),

        # Bank
        "repayment_mode": "AUTO DEBIT",
        "account_no": extracted.get("account_number"),
        "micr": "",
        "ifsc": extracted.get("ifsc_code"),
        "bank_name": extracted.get("bank_name"),
        "account_type": (
            account_type_override
            or extracted.get("account_type")
            or "31"
        ),

        # KYC
        "id_proof_no": "",
        "id_expiry": "",
        "id_doc_type": "",
        "address_proof_no": "",
        "address_expiry": "",
        "address_doc_type": "",

        # SC/ST Flag — FIX: 'Y'/'N' to match template column AO
        "sc_st_flag": sc_st_flag,

        # Property
        "plot_size": plot_size,
        "asset_cost": asset_cost,
        "amt_financed": amt_financed,

        # Notes
        "notes": extracted.get("notes"),
    }

    return row