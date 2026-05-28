import base64
import mimetypes
import os
import re
from pathlib import Path
from typing import Any, List, Optional

from openai import OpenAI
from openai.types.chat import ChatCompletionUserMessageParam, ChatCompletionSystemMessageParam
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
    micr_code: Optional[str] = None

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
You are an OCR extraction system for YEIDA RPS-10/2026 plot loan application forms (ICICI Bank).

Your job is to read every field visible in the uploaded form image(s) and return structured data.

=== FIELD EXTRACTION GUIDE ===

application_no / form_no:
  Look for "Application No", "Form No", "RPS No", or any 12-character alphanumeric code
  starting with "RPS" (e.g. RPS100132931).

applicant_name:
  Look for "Applicant Name", "Name of Applicant". Return full name as written.

plot_size:
  Look for "Plot Size", "Area", "Sq. Mtr", "Plot Area".
  Valid values are ONLY: 162, 183, 184, 200, 223, 290 (square metres).
  Return as integer. If not one of these values, return null.

category:
  Look for "Category" checkboxes or filled fields.
  Return exactly as written: GENERAL AND OTHERS, SC, ST, OBC, or EWS.

date_of_birth:
  Look for "Date of Birth", "DOB". Return in dd/mm/yyyy format.

gender:
  MALE, FEMALE, or OTHER.

marital_status:
  MARRIED, UNMARRIED, SINGLE, or WIDOWED.

mobile_no:
  Must be 10 digits. Return digits only, no spaces or dashes.

pan_number:
  Must be exactly 10 chars: 5 letters + P + 1 letter + 4 digits + 1 letter.
  Return UPPERCASE. Never guess.

account_number:
  Return as text string preserving leading zeros. Never guess.

ifsc_code:
  Look for "IFSC", "IFSC Code" on the form AND on any attached cheque copy.
  On a cheque: IFSC is usually printed on the top-left or top-right.
  Must be exactly 11 characters (4 letters + 0 + 6 alphanumeric). Never guess.

micr_code:
  Look for "MICR", "MICR Code" on the form AND at the bottom of any attached cheque copy.
  The MICR band is the row of machine-readable digits at the very bottom of a cheque —
  it contains cheque number (6 digits) + city code (3 digits) + bank code (3 digits) + account (6 digits).
  Extract the full 9-digit MICR city+bank code, OR the full MICR line if visible.
  Must be numeric only. Never guess.

bank_name:
  Look for "Bank Name" on the form AND the bank name printed on the cheque.
  Return the full official bank name in UPPERCASE (e.g. ICICI BANK, HDFC BANK, STATE BANK OF INDIA).

receipt_no / handwritten_receipt_no / printed_receipt_no:
  handwritten_receipt_no: handwritten or stamped number (prefer this).
  printed_receipt_no: pre-printed on the form.
  Return digits only (may start with 0).

sol_id:
  Look for "SOL ID", "Branch Code" - typically 4 digits.

mailing_address_1, mailing_address_2, mailing_address_3:
  Split the mailing/communication address across these three fields.

mailing_city:
  City from the mailing address.

mailing_pincode:
  6-digit PIN code from the mailing address.

employment_type:
  SALARIED, SELF-EMPLOYED, or BUSINESS.

=== RULES ===
- NEVER guess or infer PAN, account numbers, IFSC, or mobile numbers.
- If a field is not clearly visible, return null.
- Return all text fields in UPPERCASE.
- plot_size MUST be one of: 162, 183, 184, 200, 223, 290 — null otherwise.
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
        timeout=120.0,
    )

    model = model or os.environ.get(
        "OPENAI_MODEL",
        "gpt-4.1-mini",
    )

    # Build user content: text instruction + one image block per page
    # Uses Chat Completions format (type: text / image_url)
    # which is stable and supports structured Pydantic output via beta.parse
    user_content: List[dict] = [
        {
            "type": "text",
            "text": "Extract applicant data from these YEIDA RPS form images.",
        }
    ]

    for path in image_paths:
        user_content.append({
            "type": "image_url",
            "image_url": {
                "url":    encode_image_to_data_url(path),
                "detail": "high",
            },
        })

    response = client.beta.chat.completions.parse(
        model=model,
        messages=[
            {
                "role": "system",
                "content": SYSTEM_PROMPT,
            },
            {
                "role": "user",
                "content": user_content,  # type: ignore
            },
        ],
        response_format=YEIDAExtract,
    )

    parsed = response.choices[0].message.parsed

    if not parsed:
        raise RuntimeError("OCR extraction failed — model returned no structured data")

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
# REPAYMENT MODE
# =========================================================

# ICICI Bank identifiers — any of these in the bank name triggers AUTO DEBIT
_ICICI_KEYWORDS = {"ICICI"}


def resolve_repayment_mode(bank_name: Optional[str]) -> str:
    """Return AUTO DEBIT for ICICI Bank, NACH for all other banks."""
    if not bank_name:
        return "NACH"  # unknown bank — default to NACH (safer)
    if any(kw in bank_name.strip().upper() for kw in _ICICI_KEYWORDS):
        return "AUTO DEBIT"
    return "NACH"


# =========================================================
# BUSINESS RULES
# =========================================================

def apply_business_rules(
    data: YEIDAExtract,
    account_type_override=None,
    category_override=None,
    receipt_override=None,
    sol_id_override=None,
    plot_size_override=None,
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

    # Plot size — form override wins over OCR-extracted value
    raw_plot = plot_size_override or extracted.get("plot_size")
    try:
        plot_size = int(raw_plot) if raw_plot is not None else None
    except Exception:
        plot_size = None

    # Validate: must be a known size, otherwise null (no pricing lookup)
    if plot_size not in PRICING:
        plot_size = None

    # Pricing — auto-filled from PRICING table based on plot_size
    if plot_size is not None:
        asset_cost   = PRICING[plot_size]["asset_cost"]
        amt_financed = PRICING[plot_size]["amt_financed"]
    else:
        asset_cost   = None
        amt_financed = None

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
        # AUTO DEBIT for ICICI Bank, NACH for all other banks
        "repayment_mode": resolve_repayment_mode(extracted.get("bank_name")),
        "account_no": extracted.get("account_number"),
        "micr": extracted.get("micr_code") or "",
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