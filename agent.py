import base64
import json
import mimetypes
import os
import re
from pathlib import Path
from typing import List, Optional

from openai import OpenAI
from pydantic import BaseModel, Field


class YEIDAExtract(BaseModel):
    # Core application details
    form_no: Optional[str] = Field(default=None, description="YEIDA form number / RPS number")
    application_no: Optional[str] = Field(default=None, description="YEIDA application number / RPS number")
    applicant_name: Optional[str] = None
    father_or_husband_name: Optional[str] = None
    category: Optional[str] = Field(default=None, description="GENERAL, SC, ST, OBC, etc.")
    sub_category: Optional[str] = None
    plot_size: Optional[int] = Field(default=None, description="Plot size in square metres")
    date_of_birth: Optional[str] = Field(default=None, description="DD/MM/YYYY if visible")
    gender: Optional[str] = None
    marital_status: Optional[str] = None
    age: Optional[int] = None
    handicapped: Optional[str] = None
    payment_plan: Optional[str] = None
    registration_money: Optional[float] = None
    form_fees: Optional[float] = None
    deposit_amount: Optional[float] = None
    paid_by: Optional[str] = None

    # Co-applicant
    co_applicant_name: Optional[str] = None
    relation_with_applicant: Optional[str] = None
    co_father_or_husband_name: Optional[str] = None

    # Address/contact
    mailing_address_1: Optional[str] = None
    mailing_address_2: Optional[str] = None
    mailing_address_3: Optional[str] = None
    mailing_city: Optional[str] = None
    mailing_pincode: Optional[str] = None
    permanent_address_1: Optional[str] = None
    permanent_address_2: Optional[str] = None
    permanent_address_3: Optional[str] = None
    permanent_city: Optional[str] = None
    permanent_pincode: Optional[str] = None
    mobile_no: Optional[str] = None
    email: Optional[str] = None

    # Bank details
    account_name: Optional[str] = None
    bank_name: Optional[str] = None
    branch_name: Optional[str] = None
    account_number: Optional[str] = None
    ifsc_code: Optional[str] = None
    pan_number: Optional[str] = None
    id_proof_no: Optional[str] = Field(default=None, description="Only the last digit of Aadhaar/id proof if visible; otherwise null")

    # Receipt / handwritten cover details
    printed_receipt_no: Optional[str] = None
    handwritten_receipt_no: Optional[str] = Field(default=None, description="Blue-circled or handwritten receipt number. Prefer this over printed receipt.")
    receipt_no: Optional[str] = Field(default=None, description="Final receipt number to use. Prefer handwritten/blue-circled number.")
    sol_id: Optional[str] = None
    emd_branch: Optional[str] = None
    payment_date: Optional[str] = None
    dd_utr_no: Optional[str] = None
    payment_id: Optional[str] = None

    # Overrides/rules
    account_type: Optional[str] = None
    employment_type: Optional[str] = None

    # Pricing derived by code, but model can extract if visible
    asset_cost: Optional[float] = None
    customer_contribution: Optional[float] = None
    amt_financed: Optional[float] = Field(default=None, description="Bank loan / amount financed")

    notes: Optional[str] = None


SYSTEM_PROMPT = """
You are a precise data-extraction agent for YEIDA / Yamuna Expressway Industrial Development Authority RPS-10 2026 forms and EMD funding/loan agreement packets.

Extract details from ALL uploaded images for ONE applicant. Images may include:
- YEIDA confirmation/application form
- mailing/permanent address page
- bank details page
- online payment receipt
- handwritten cover page with receipt number / account type / SOL ID

Important rules:
1. Never invent Aadhaar, PAN, bank account, DOB, mobile, receipt, or address data. Use null if not visible.
2. For ID proof/Aadhaar number, return ONLY the LAST DIGIT if visible. If Aadhaar is not visible, return null.
3. If a handwritten or blue-circled receipt number appears, use that as handwritten_receipt_no and final receipt_no, even if the printed receipt page shows a different online receipt number.
4. If no handwritten/circled receipt number appears, use printed receipt number as receipt_no.
5. Extract account type only if visible/handwritten, otherwise leave null. The application will apply default/override.
6. Employment type should be SALARIED unless explicitly contradicted.
7. Normalize categories: GENERAL, SC, ST, OBC, EWS, GENERAL AND OTHERS.
8. Mobile number must be digits only, normally 10 digits. If multiple numbers appear, prefer the number on receipt/form for the applicant.
9. Do not include explanations outside the schema.
""".strip()


def encode_image_to_data_url(path: Path) -> str:
    mime, _ = mimetypes.guess_type(path)
    if not mime:
        mime = "image/jpeg"
    encoded = base64.b64encode(path.read_bytes()).decode("utf-8")
    return f"data:{mime};base64,{encoded}"


def extract_from_images(image_paths: List[Path], model: Optional[str] = None) -> YEIDAExtract:
    if not image_paths:
        raise ValueError("No image files were provided.")

    client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))
    model = model or os.environ.get("OPENAI_MODEL", "gpt-4.1-mini")

    content = [
        {
            "type": "input_text",
            "text": (
                "Extract one YEIDA applicant entry from these images. "
                "Prefer handwritten/blue-circled receipt number for final receipt_no. "
                "Return structured data only."
            ),
        }
    ]

    for p in image_paths:
        content.append({"type": "input_image", "image_url": encode_image_to_data_url(p)})

    # Responses API with SDK parsing into the Pydantic schema.
    response = client.responses.parse(
        model=model,
        input=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": content},
        ],
        text_format=YEIDAExtract,
    )

    parsed = response.output_parsed
    if parsed is None:
        raise RuntimeError("Model did not return parsed structured data.")
    return parsed


def only_digits(value: Optional[str]) -> Optional[str]:
    if value is None:
        return None
    digits = re.sub(r"\D", "", str(value))
    return digits or None


def normalize_text(value: Optional[str]) -> Optional[str]:
    if value is None:
        return None
    text = re.sub(r"\s+", " ", str(value)).strip()
    return text or None


def normalize_category(value: Optional[str]) -> Optional[str]:
    text = normalize_text(value)
    if not text:
        return None
    upper = text.upper()
    if "SC" == upper or "SCHEDULED CASTE" in upper:
        return "SC"
    if "ST" == upper or "SCHEDULED TRIBE" in upper:
        return "ST"
    if "OBC" in upper:
        return "OBC"
    if "EWS" in upper:
        return "EWS"
    if "GENERAL" in upper:
        return "GENERAL AND OTHERS"
    return upper


# Pricing table from the YEIDA EMD finance sheet used in the uploaded documents.
# Asset cost is registration/earnest money * 10. Amt financed is bank loan.
PRICING = {
    "GENERAL": {
        162: {"registration_money": 587412, "deposit_amount": 58741, "amt_financed": 528671, "asset_cost": 5874120},
        183: {"registration_money": 663588, "deposit_amount": 66359, "amt_financed": 597229, "asset_cost": 6635880},
        184: {"registration_money": 667184, "deposit_amount": 66718, "amt_financed": 600466, "asset_cost": 6671840},
        200: {"registration_money": 725200, "deposit_amount": 72520, "amt_financed": 652680, "asset_cost": 7252000},
        223: {"registration_money": 808598, "deposit_amount": 80860, "amt_financed": 727738, "asset_cost": 8085980},
        290: {"registration_money": 1051540, "deposit_amount": 105154, "amt_financed": 946386, "asset_cost": 10515400},
    },
    "SCST": {
        162: {"registration_money": 293706, "deposit_amount": 29371, "amt_financed": 264335, "asset_cost": 2937060},
        183: {"registration_money": 331779, "deposit_amount": 33178, "amt_financed": 298601, "asset_cost": 3317790},
        184: {"registration_money": 333592, "deposit_amount": 33359, "amt_financed": 300233, "asset_cost": 3335920},
        200: {"registration_money": 362600, "deposit_amount": 36260, "amt_financed": 326340, "asset_cost": 3626000},
        223: {"registration_money": 404299, "deposit_amount": 40430, "amt_financed": 363869, "asset_cost": 4042990},
        290: {"registration_money": 525770, "deposit_amount": 52577, "amt_financed": 473193, "asset_cost": 5257700},
    },
}


def pricing_key(category: Optional[str]) -> str:
    cat = normalize_category(category) or "GENERAL"
    return "SCST" if cat in {"SC", "ST"} else "GENERAL"


def apply_business_rules(
    data: YEIDAExtract,
    account_type_override: Optional[str] = None,
    category_override: Optional[str] = None,
    receipt_override: Optional[str] = None,
    sol_id_override: Optional[str] = None,
) -> dict:
    row = data.model_dump()

    # Normalize text fields
    for k, v in list(row.items()):
        if isinstance(v, str):
            row[k] = normalize_text(v)

    # Application/Form no fallback
    row["application_no"] = row.get("application_no") or row.get("form_no")
    row["form_no"] = row.get("form_no") or row.get("application_no")

    # Receipt priority: manual override > handwritten/blue circle > model final > printed
    final_receipt = (
        only_digits(receipt_override)
        or only_digits(row.get("handwritten_receipt_no"))
        or only_digits(row.get("receipt_no"))
        or only_digits(row.get("printed_receipt_no"))
    )
    row["receipt_no"] = final_receipt

    # Account type priority: manual override > extracted/handwritten > default
    row["account_type"] = normalize_text(account_type_override) or normalize_text(row.get("account_type")) or os.environ.get("DEFAULT_ACCOUNT_TYPE", "31")

    # Employment always salaried unless you change default in .env
    row["employment_type"] = os.environ.get("DEFAULT_EMPLOYMENT", "SALARIED")

    # Category override + normalization
    row["category"] = normalize_category(category_override or row.get("category"))
    row["sub_category"] = normalize_category(row.get("sub_category"))

    # IDs / contact cleanup
    row["mobile_no"] = only_digits(row.get("mobile_no"))
    row["id_proof_no"] = only_digits(row.get("id_proof_no"))[-1:] if only_digits(row.get("id_proof_no")) else None
    row["pan_number"] = normalize_text(row.get("pan_number"))
    row["ifsc_code"] = normalize_text(row.get("ifsc_code"))
    row["sol_id"] = only_digits(sol_id_override) or only_digits(row.get("sol_id"))

    # Pricing by category + plot size. Prefer calculated standard values because receipt may show contribution only.
    plot_size = row.get("plot_size")
    try:
        plot_size = int(plot_size) if plot_size is not None else None
        row["plot_size"] = plot_size
    except Exception:
        plot_size = None

    if plot_size:
        table = PRICING.get(pricing_key(row.get("category")), {})
        price = table.get(plot_size)
        if price:
            row["registration_money"] = float(price["registration_money"])
            row["deposit_amount"] = float(price["deposit_amount"])
            row["customer_contribution"] = float(price["deposit_amount"])
            row["amt_financed"] = float(price["amt_financed"])
            row["asset_cost"] = float(price["asset_cost"])

    # Form fee defaults
    if row.get("form_fees") is None:
        row["form_fees"] = 600.0

    # Paid by defaults
    if not row.get("paid_by"):
        row["paid_by"] = "EMD"

    return row
