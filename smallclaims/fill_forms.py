"""
Oakland Encampment Small Claims Form Autofiller
Fills California court forms from a case JSON file.
Zero LLM calls — deterministic pypdf field filling.

Usage:
    python fill_forms.py cases/jane_doe.json      # fill one case
    python fill_forms.py cases/                   # fill all cases in directory
    python fill_forms.py --new cases/new.json     # generate blank case template
"""

import json
import io
import sys
import os
import re
import argparse
from datetime import datetime
from pathlib import Path

try:
    from pypdf import PdfReader, PdfWriter
except Exception:
    # Compatibility fallback for environments that still provide PyPDF2.
    from PyPDF2 import PdfReader, PdfWriter  # type: ignore

try:
    from pypdf.generic import NameObject, BooleanObject, DictionaryObject
except Exception:
    NameObject = BooleanObject = DictionaryObject = None

try:
    from pypdf.generic import TextStringObject
except Exception:
    TextStringObject = None


def _safe_print(*args, **kwargs):
    """Best-effort logging that never breaks form generation."""
    try:
        print(*args, **kwargs)
    except Exception:
        pass

# Optional: flatten PDFs to ensure appearances are rendered in all viewers.
def _flatten_pdf(path: str, scale: float = 2.0) -> None:
    """Rasterize each page and re-save as a PDF to bake in appearances.

    This uses PyMuPDF (package name `pymupdf`). It preserves visual fidelity
    at the cost of producing a rasterized PDF (larger file, not editable).
    If PyMuPDF is not available, this is a no-op.
    """
    try:
        import fitz  # PyMuPDF
    except Exception:
        return

    try:
        # Build an image-only PDF using reportlab to avoid carrying over any
        # AcroForm or widget annotations from the original PDF.
        doc = fitz.open(path)
        mat = fitz.Matrix(scale, scale)
        try:
            from reportlab.pdfgen.canvas import Canvas
            from reportlab.lib.utils import ImageReader
            import io as _io

            buf = _io.BytesIO()
            c = Canvas(buf)
            for page in doc:
                pix = page.get_pixmap(matrix=mat)
                png = pix.tobytes("png")
                img = ImageReader(_io.BytesIO(png))
                # Use pixel dimensions for page size so the image fills the page
                w, h = pix.width, pix.height
                c.setPageSize((w, h))
                c.drawImage(img, 0, 0, width=w, height=h)
                c.showPage()
            c.save()
            tmp = path + ".flattmp"
            with open(tmp, 'wb') as f:
                f.write(buf.getvalue())
            buf.close()
            doc.close()
            os.replace(tmp, path)
        except Exception:
            # If reportlab isn't available or fails, fall back to PyMuPDF insertion
            new = fitz.open()
            for page in doc:
                pix = page.get_pixmap(matrix=mat)
                img_pdf = fitz.open("pdf", pix.tobytes("pdf"))
                new.insert_pdf(img_pdf)
            tmp = path + ".flattmp"
            new.save(tmp)
            new.close()
            doc.close()
            os.replace(tmp, path)
    except Exception:
        # Non-fatal: if flattening fails, leave original PDF as-is
        return


# ─────────────────────────────────────────────────────────────
# COURT INFO — dynamic, reads from case["court"]
# ─────────────────────────────────────────────────────────────

_DEFAULT_COURT = {
    "county":  "Alameda",
    "address": "1225 Fallon Street",
    "city":    "Oakland",
    "zip":     "94612",
}


def _court_info(case: dict) -> str:
    ct = case.get("court", _DEFAULT_COURT)
    county  = ct.get("county",  _DEFAULT_COURT["county"])
    address = ct.get("address", _DEFAULT_COURT["address"])
    city    = ct.get("city",    _DEFAULT_COURT["city"])
    zip_    = ct.get("zip",     _DEFAULT_COURT["zip"])
    return (
        f"Superior Court of California, County of {county}\n"
        f"{address}\n"
        f"{city}, CA {zip_}"
    )


def _venue_zip(case: dict) -> str:
    defendant = case.get("defendant") or {}
    if str(defendant.get("zip") or "").strip():
        return str(defendant.get("zip")).strip()
    return case.get("court", _DEFAULT_COURT).get("zip", _DEFAULT_COURT["zip"])


DEFENDANT_DEFAULTS = {
    "city_of_oakland": {
        "name": "City of Oakland",
        "address": "One Frank H. Ogawa Plaza",
        "city": "Oakland",
        "state": "CA",
        "zip": "94612",
        "agent_name": "City Clerk",
        "agent_title": "City Clerk",
        "agent_address": "One Frank H. Ogawa Plaza",
        "agent_city": "Oakland",
        "agent_state": "CA",
        "agent_zip": "94612",
    }
}


def _forms_generated_date(case: dict | None = None) -> str:
    """Return the generation date in MM/DD/YYYY format."""
    stamp = str((case or {}).get("forms_generated_at") or "").strip()
    if stamp:
        try:
            return datetime.fromisoformat(stamp.replace("Z", "+00:00")).strftime("%m/%d/%Y")
        except Exception:
            pass
    return datetime.now().strftime("%m/%d/%Y")



# ─────────────────────────────────────────────────────────────
# FIELD METADATA (checkbox on/off values)
# ─────────────────────────────────────────────────────────────

def load_field_meta(json_path):
    with open(json_path) as f:
        fields = json.load(f)
    return {item["field_id"]: item for item in fields}


def _checkbox_value(meta, field_id, checked):
    if checked:
        return meta.get(field_id, {}).get("checked_value", "/Yes")
    return meta.get(field_id, {}).get("unchecked_value", "/Off")


def _caption_fields(case, court_key=None, case_number_key=None, case_name_key=None):
    values = {
    }
    if court_key:
        values[court_key] = _court_info(case)
    if case_number_key:
        values[case_number_key] = case.get("case_number", "")
    if case_name_key:
        values[case_name_key] = _case_name(case)
    return values


def _plaintiff_contact(case: dict) -> dict:
    """Normalize plaintiff contact fields used across form headers."""
    p = case.get("plaintiff", {}) or {}
    state = p.get("state", "CA")
    zip_code = p.get("zip", "")
    mailing_address = ", ".join(
        part for part in [
            p.get("street", ""),
            p.get("city", ""),
            f"{state} {zip_code}".strip(),
        ] if part
    )
    return {
        "name": p.get("name", ""),
        "street": p.get("street", ""),
        "city": p.get("city", ""),
        "state": state,
        "zip": zip_code,
        "phone": p.get("phone", ""),
        "email": p.get("email", ""),
        "mailing_address": mailing_address,
    }


def _plaintiff_fields(case: dict, field_map: dict[str, str]) -> dict:
    """Map standardized plaintiff contact keys onto concrete PDF field ids."""
    contact = _plaintiff_contact(case)
    return {pdf_field: contact.get(contact_key, "") for pdf_field, contact_key in field_map.items()}


# ─────────────────────────────────────────────────────────────
# VALIDATION
# ─────────────────────────────────────────────────────────────

def validate_case(case):
    """Raise ValueError with clear message if required fields are missing."""
    errors = []

    def _is_public_entity(defendant: dict) -> bool:
        name = str((defendant or {}).get("name") or "").strip().lower()
        if not name:
            return False
        public_tokens = (
            "city", "county", "department", "agency", "district",
            "state", "public", "authority", "board", "university",
        )
        return any(tok in name for tok in public_tokens)

    p = case.get("plaintiff", {})
    if not p.get("name"):
        errors.append("plaintiff.name is required")
    if not p.get("street") and not p.get("city"):
        errors.append(
            "plaintiff needs at least street or city "
            "(use 'c/o <address>' for unhoused clients)"
        )

    claim = case.get("claim", {})
    if not claim.get("amount"):
        errors.append("claim.amount is required")
    if not claim.get("reason"):
        errors.append("claim.reason is required")
    if not claim.get("incident_date"):
        errors.append("claim.incident_date is required (format: MM/DD/YYYY)")
    defendant = case.get("defendant") or {}
    if _is_public_entity(defendant) and not claim.get("govt_claim_filed_date"):
        errors.append(
            "claim.govt_claim_filed_date is required when suing a public entity"
        )

    fee_waiver = case.get("fee_waiver") or {}
    basis = str(fee_waiver.get("basis") or "").strip().lower()
    if basis == "5a":
        has_public_benefit = any([
            bool(fee_waiver.get("receives_medi_cal")),
            bool(fee_waiver.get("receives_snap")),
            bool(fee_waiver.get("receives_calworks")),
        ])
        if not has_public_benefit:
            errors.append(
                "fee_waiver requires at least one selected public benefit when basis is 5a"
            )

    if errors:
        raise ValueError(
            "Case validation failed:\n" + "\n".join(f"  • {e}" for e in errors)
        )


# ─────────────────────────────────────────────────────────────
# CORE PDF WRITE
# ─────────────────────────────────────────────────────────────

def _write_pdf(template_path, output_path, values, field_appearances: dict[str, str] | None = None):
    # Keep template bytes in memory for the full write lifecycle.
    # Some runtimes can invalidate file-backed handles while pypdf still
    # lazily reads from the source document during writer.write().
    template_stream = io.BytesIO(Path(template_path).read_bytes())
    reader = PdfReader(template_stream)
    writer = PdfWriter()
    writer.append(reader)

    for page in writer.pages:
        writer.update_page_form_field_values(page, values)

    if field_appearances:
        try:
            for page in writer.pages:
                annots = page.get("/Annots") or []
                for annot_ref in annots:
                    annot = annot_ref.get_object()
                    names = {str(annot.get("/T") or "")}
                    parent = annot.get("/Parent")
                    parent_obj = parent.get_object() if parent else None
                    if parent_obj is not None:
                        names.add(str(parent_obj.get("/T") or ""))
                    for field_name, appearance in field_appearances.items():
                        if field_name not in names:
                            continue
                        da_value = TextStringObject(appearance) if TextStringObject else appearance
                        annot.update({NameObject("/DA"): da_value})
                        if parent_obj is not None:
                            parent_obj.update({NameObject("/DA"): da_value})
        except Exception:
            pass

    # Ensure PDF viewers render the updated field values by setting
    # the AcroForm /NeedAppearances flag. Some viewers require this
    # to generate appearance streams for filled fields when the PDF
    # is opened or saved elsewhere.
    try:
        root = writer._root_object
        # Set the AcroForm /NeedAppearances flag so PDF viewers render filled fields.
        # Use generic PDF object wrappers when available, otherwise fall back to
        # plain keys for broader compatibility across pypdf/PyPDF2 variants.
        if NameObject and BooleanObject and DictionaryObject:
            try:
                acro = root.get(NameObject("/AcroForm")) if hasattr(root, "get") else None
            except Exception:
                acro = None

            if acro is None:
                root.update({
                    NameObject("/AcroForm"): DictionaryObject({
                        NameObject("/NeedAppearances"): BooleanObject(True)
                    })
                })
            else:
                try:
                    acro.update({NameObject("/NeedAppearances"): BooleanObject(True)})
                except Exception:
                    root.update({
                        NameObject("/AcroForm"): DictionaryObject({
                            NameObject("/NeedAppearances"): BooleanObject(True)
                        })
                    })
        else:
            acro = root.get("/AcroForm") if hasattr(root, "get") else None
            if acro is not None and hasattr(acro, "update"):
                try:
                    acro.update({"/NeedAppearances": True})
                except Exception:
                    pass
    except Exception:
        # Non-fatal: continue writing even if we can't set the flag.
        pass

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    with open(output_path, "wb") as f:
        writer.write(f)
    template_stream.close()

    # Try to flatten the PDF so appearances are visually baked in for all viewers.
    # Flattening is optional and will be skipped if PyMuPDF (`pymupdf`) isn't installed.
    try:
        _flatten_pdf(output_path)
    except Exception:
        pass


def _case_name(case):
    d = case.get("defendant", DEFENDANT_DEFAULTS["city_of_oakland"])
    defendant_name = d.get("name", "City of Oakland")
    if case.get("additional_defendants"):
        defendant_name = f"{defendant_name}, et al."
    return f"{case['plaintiff']['name']} v. {defendant_name}"


def _ellipsize_middle(text: str, max_len: int) -> str:
    text = str(text or "").strip()
    if len(text) <= max_len:
        return text
    if max_len <= 3:
        return text[:max_len]
    keep_left = (max_len - 3) // 2
    keep_right = max_len - 3 - keep_left
    return f"{text[:keep_left]}...{text[-keep_right:]}"


def _fw003_case_name(case: dict) -> str:
    plaintiff_name = str((case.get("plaintiff") or {}).get("name") or "Plaintiff").strip()
    defendant = case.get("defendant") or DEFENDANT_DEFAULTS["city_of_oakland"]
    defendant_name = str(defendant.get("name") or "City of Oakland").strip()
    if case.get("additional_defendants"):
        defendant_name = f"{defendant_name}, et al."

    plaintiff_name = _ellipsize_middle(plaintiff_name, 28)
    defendant_name = _ellipsize_middle(defendant_name, 28)
    return f"{plaintiff_name} v. {defendant_name}"


def _sc100_section5_fields(case: dict, meta: dict) -> dict:
    """Build SC-100 page 3 section 5 fields (a-e), defaulting to option a."""
    claim = case.get("claim") or {}
    selected = str(claim.get("sc100_section5_reason") or "a").strip().lower()
    if selected not in {"a", "b", "c", "d", "e"}:
        selected = "a"

    option_to_field = {
        "a": "SC-100[0].Page3[0].List5[0].Lia[0].Checkbox5cb[0]",
        "b": "SC-100[0].Page3[0].List5[0].Lib[0].Checkbox5cb[0]",
        "c": "SC-100[0].Page3[0].List5[0].Lic[0].Checkbox5cb[0]",
        "d": "SC-100[0].Page3[0].List5[0].Lid[0].Checkbox5cb[0]",
        "e": "SC-100[0].Page3[0].List5[0].Lie[0].Checkbox5cb[0]",
    }

    values = {
        field_id: _checkbox_value(meta, field_id, option == selected)
        for option, field_id in option_to_field.items()
    }
    values["SC-100[0].Page3[0].List5[0].Lie[0].FillField55[0]"] = str(
        claim.get("sc100_section5_other") or ""
    ).strip()
    return values


# ─────────────────────────────────────────────────────────────
# SC-100  Plaintiff's Claim and ORDER to Go to Small Claims Court
# ─────────────────────────────────────────────────────────────

def fill_sc100(case, template_path, output_path, field_meta_path):
    meta = load_field_meta(field_meta_path)
    contact = _plaintiff_contact(case)
    d = case.get("defendant", DEFENDANT_DEFAULTS["city_of_oakland"])
    claim = case["claim"]
    filing = case.get("filing", {})
    generated_date = _forms_generated_date(case)

    values = {
        # Header
        "SC-100[0].Page1[0].CaptionRight[0].County[0].CourtInfo[0]": _court_info(case),

        # Plaintiff
        **_plaintiff_fields(case, {
            "SC-100[0].Page2[0].List1[0].Item1[0].PlaintiffName1[0]": "name",
            "SC-100[0].Page2[0].List1[0].Item1[0].PlaintiffAddress1[0]": "street",
            "SC-100[0].Page2[0].List1[0].Item1[0].PlaintiffCity1[0]": "city",
            "SC-100[0].Page2[0].List1[0].Item1[0].PlaintiffState1[0]": "state",
            "SC-100[0].Page2[0].List1[0].Item1[0].PlaintiffZip1[0]": "zip",
            "SC-100[0].Page2[0].List1[0].Item1[0].PlaintiffPhone1[0]": "phone",
            "SC-100[0].Page2[0].List1[0].Item1[0].EmailAdd1[0]": "email",
        }),

        # Defendant
        "SC-100[0].Page2[0].List2[0].item2[0].DefendantName1[0]":    d.get("name", "City of Oakland"),
        "SC-100[0].Page2[0].List2[0].item2[0].DefendantAddress1[0]": d.get("address", ""),
        "SC-100[0].Page2[0].List2[0].item2[0].DefendantCity1[0]":    d.get("city", "Oakland"),
        "SC-100[0].Page2[0].List2[0].item2[0].DefendantState1[0]":   d.get("state", "CA"),
        "SC-100[0].Page2[0].List2[0].item2[0].DefendantZip1[0]":     d.get("zip", "94612"),
        "SC-100[0].Page2[0].List2[0].item2[0].DefendantJob1[0]":     d.get("agent_name", "City Clerk"),
        "SC-100[0].Page2[0].List2[0].item2[0].DefendantAddress2[0]": d.get("agent_address", ""),
        "SC-100[0].Page2[0].List2[0].item2[0].DefendantCity2[0]":    d.get("agent_city", "Oakland"),
        "SC-100[0].Page2[0].List2[0].item2[0].DefendantState2[0]":   d.get("agent_state", "CA"),
        "SC-100[0].Page2[0].List2[0].item2[0].DefendantZip2[0]":     d.get("agent_zip", "94612"),

        # Claim amount + reason
        "SC-100[0].Page2[0].List3[0].PlaintiffClaimAmount1[0]": str(claim["amount"]),
        "SC-100[0].Page2[0].List3[0].Lia[0].FillField2[0]":    claim["reason"],

        # When did this happen
        "SC-100[0].Page3[0].List3[0].Lib[0].Date1[0]": claim.get("incident_date", ""),
        "SC-100[0].Page3[0].List3[0].Lib[0].Date2[0]": claim.get("date_started", ""),
        "SC-100[0].Page3[0].List3[0].Lib[0].Date3[0]": claim.get("date_through", ""),

        # How damages calculated
        "SC-100[0].Page3[0].List3[0].Lic[0].FillField1[0]": claim.get("damages_calculation", ""),

        # Have you asked defendant to pay?
        "SC-100[0].Page3[0].List4[0].Item4[0].Checkbox50[0]": _checkbox_value(meta,
            "SC-100[0].Page3[0].List4[0].Item4[0].Checkbox50[0]",
            filing.get("demanded_payment", True),
        ),
        "SC-100[0].Page3[0].List4[0].Item4[0].Checkbox50[1]": _checkbox_value(meta,
            "SC-100[0].Page3[0].List4[0].Item4[0].Checkbox50[1]",
            not filing.get("demanded_payment", True),
        ),

        # Section 5 — why this lawsuit is being filed
        **_sc100_section5_fields(case, meta),

        # Section 6 zip — default to first defendant zip
        "SC-100[0].Page3[0].List6[0].item6[0].ZipCode1[0]": _venue_zip(case),

        # Attorney-client fee dispute? No
        "SC-100[0].Page3[0].List7[0].item7[0].Checkbox60[0]": _checkbox_value(meta,
            "SC-100[0].Page3[0].List7[0].item7[0].Checkbox60[0]", False
        ),
        "SC-100[0].Page3[0].List7[0].item7[0].Checkbox60[1]": _checkbox_value(meta,
            "SC-100[0].Page3[0].List7[0].item7[0].Checkbox60[1]", True
        ),

        # Suing a public entity? Yes + claim date
        "SC-100[0].Page3[0].List8[0].item8[0].Checkbox61[0]": _checkbox_value(meta,
            "SC-100[0].Page3[0].List8[0].item8[0].Checkbox61[0]", True
        ),
        "SC-100[0].Page3[0].List8[0].item8[0].Checkbox61[1]": _checkbox_value(meta,
            "SC-100[0].Page3[0].List8[0].item8[0].Checkbox61[1]", False
        ),
        "SC-100[0].Page3[0].List8[0].item8[0].Date4[0]": claim.get("govt_claim_filed_date", ""),

        # Filed more than 12 claims this year? No
        "SC-100[0].Page4[0].List9[0].Item9[0].Checkbox62[0]": _checkbox_value(meta,
            "SC-100[0].Page4[0].List9[0].Item9[0].Checkbox62[0]", False
        ),
        "SC-100[0].Page4[0].List9[0].Item9[0].Checkbox62[1]": _checkbox_value(meta,
            "SC-100[0].Page4[0].List9[0].Item9[0].Checkbox62[1]", True
        ),

        # Claim for more than $2,500? Yes
        "SC-100[0].Page4[0].List10[0].li10[0].Checkbox63[0]": _checkbox_value(meta,
            "SC-100[0].Page4[0].List10[0].li10[0].Checkbox63[0]", True
        ),
        "SC-100[0].Page4[0].List10[0].li10[0].Checkbox63[1]": _checkbox_value(meta,
            "SC-100[0].Page4[0].List10[0].li10[0].Checkbox63[1]", False
        ),

        # Signature
        "SC-100[0].Page4[0].Sign[0].Date1[0]":          generated_date,
        "SC-100[0].Page4[0].Sign[0].PlaintiffName1[0]":  contact["name"],

        # Repeated caption fields
        **{
            f"SC-100[0].Page{page}[0].PxCaption[0].Plaintiff[0]": contact["name"]
            for page in (2, 3, 4)
        },
    }

    _write_pdf(template_path, output_path, values)
    try:
        _flatten_pdf(output_path)
    except Exception:
        pass
    _safe_print(f"  ✓ SC-100  → {output_path}")


# ─────────────────────────────────────────────────────────────
# FW-001  Request to Waive Court Fees
# ─────────────────────────────────────────────────────────────

def fill_fw001(case, template_path, output_path, field_meta_path):
    meta = load_field_meta(field_meta_path)
    contact = _plaintiff_contact(case)
    fw = case.get("fee_waiver", {})
    generated_date = _forms_generated_date(case)

    basis = fw.get("basis", "5c")  # "5a" | "5b" | "5c"

    values = {
        **_caption_fields(
            case,
            "FW-001[0].Page1[0].RightCaption[0].CourtInfo[0]",
            "FW-001[0].Page1[0].RightCaption[0].CaseName[0]",
        ),

        # Section 1: Petitioner info
        **_plaintiff_fields(case, {
            "FW-001[0].Page1[0].List1[0].item1[0].PetitionerName1[0]": "name",
            "FW-001[0].Page1[0].List1[0].item1[0].PetitionerStrAddress[0]": "street",
            "FW-001[0].Page1[0].List1[0].item1[0].PetitionerCity[0]": "city",
            "FW-001[0].Page1[0].List1[0].item1[0].PetitionerState[0]": "state",
            "FW-001[0].Page1[0].List1[0].item1[0].PetitionerZip[0]": "zip",
            "FW-001[0].Page1[0].List1[0].item1[0].PetitionerTel[0]": "phone",
        }),

        # Section 4: Superior Court fees
        "FW-001[0].Page1[0].List4[0].item4[0].WaiveSuperiorCrtFee[0]": _checkbox_value(meta,
            "FW-001[0].Page1[0].List4[0].item4[0].WaiveSuperiorCrtFee[0]", True
        ),

        # Section 5a: Public benefits
        "FW-001[0].Page1[0].List5[0].Lia[0].PublicBenefitReceived[0]": _checkbox_value(meta,
            "FW-001[0].Page1[0].List5[0].Lia[0].PublicBenefitReceived[0]", basis == "5a"
        ),
        "FW-001[0].Page1[0].List5[0].Lia[0].PublicBenefitSNAP[0]": _checkbox_value(meta,
            "FW-001[0].Page1[0].List5[0].Lia[0].PublicBenefitSNAP[0]",
            fw.get("receives_snap", False),
        ),
        "FW-001[0].Page1[0].List5[0].Lia[0].PublicBenefitMediCal[0]": _checkbox_value(meta,
            "FW-001[0].Page1[0].List5[0].Lia[0].PublicBenefitMediCal[0]",
            fw.get("receives_medi_cal", False),
        ),
        "FW-001[0].Page1[0].List5[0].Lia[0].PublicBenefitCalWORKSTANF[0]": _checkbox_value(meta,
            "FW-001[0].Page1[0].List5[0].Lia[0].PublicBenefitCalWORKSTANF[0]",
            fw.get("receives_calworks", False),
        ),

        # Section 5b: Gross income below threshold
        "FW-001[0].Page1[0].List5[0].Lib[0].GrossMonthIncomeLess[0]": _checkbox_value(meta,
            "FW-001[0].Page1[0].List5[0].Lib[0].GrossMonthIncomeLess[0]", basis == "5b"
        ),

        # Section 5c: Cannot afford fees (most common for our clients)
        "FW-001[0].Page1[0].List5[0].Lic[0].IncomeInsufficientRequest[0]": _checkbox_value(meta,
            "FW-001[0].Page1[0].List5[0].Lic[0].IncomeInsufficientRequest[0]", basis == "5c"
        ),
        "FW-001[0].Page1[0].List5[0].Lic[0].FeeRequestDef[0]": _checkbox_value(meta,
            "FW-001[0].Page1[0].List5[0].Lic[0].FeeRequestDef[0]",
            basis == "5c" and fw.get("waive_option", "all") == "all",
        ),

        # Signature
        "FW-001[0].Page1[0].Sign[0].SigDate[0]":        generated_date,
        "FW-001[0].Page1[0].Sign[0].PetitionerName[0]": contact["name"],

        # Page 2 caption + income
        "FW-001[0].Page2[0].pXCaption[0].PetitionerName1[0]":           contact["name"],
        "FW-001[0].Page2[0].List8[0].Lia[0].IncomeSource1[0]":          fw.get("income_source_1", ""),
        "FW-001[0].Page2[0].List8[0].Lia[0].IncomeAmount1[0]":          str(fw.get("income_amount_1", "")),
        "FW-001[0].Page2[0].List8[0].Lib[0].TotalIncome[0]":            str(fw.get("total_monthly_income", "")),

        # Page 2 expenses
        "FW-001[0].Page2[0].List11[0].Lib[0].ExpenseHousing[0]":        str(fw.get("expense_housing", "")),
        "FW-001[0].Page2[0].List11[0].Lic[0].ExpenseFoodSupplies[0]":   str(fw.get("expense_food", "")),
        "FW-001[0].Page2[0].List11[0].Lid[0].ExpenseUtilitiesPhone[0]": str(fw.get("expense_utilities", "")),
        "FW-001[0].Page2[0].List11[0].Lig[0].ExpenseMedicalDental[0]":  str(fw.get("expense_medical", "")),
        "FW-001[0].Page2[0].List11[0].Lik[0].ExpenseTransportation[0]": str(fw.get("expense_transport", "")),
        "FW-001[0].Page2[0].List11[0].Total[0].Totalmonthlyexpenses[0]": str(fw.get("total_monthly_expenses", "")),
    }

    _write_pdf(template_path, output_path, values)
    try:
        _flatten_pdf(output_path)
    except Exception:
        pass
    _safe_print(f"  ✓ FW-001  → {output_path}")


# ─────────────────────────────────────────────────────────────
# FW-003  Order on Court Fee Waiver  (court-completed — we pre-fill header only)
# ─────────────────────────────────────────────────────────────

def fill_fw003(case, template_path, output_path):
    contact = _plaintiff_contact(case)

    values = {
        "FW-003[0].Page1[0].Stamp_court_case[0].CourtInfo_ft[0]":  _court_info(case),
        "FW-003[0].Page1[0].Stamp_court_case[0].CaseNumber_ft[0]": case.get("case_number", ""),
        "FW-003[0].Page1[0].Stamp_court_case[0].CaseName_ft[0]":   _fw003_case_name(case),
        "FW-003[0].Page1[0].PersonWaivingName_ft[0]":              contact["name"],
        "FW-003[0].Page1[0].FillText23[0]":                        contact["street"],
        "FW-003[0].Page1[0].FillText21[0]":                        contact["city"],
        "FW-003[0].Page1[0].FillText20[0]":                        contact["state"],
        "FW-003[0].Page1[0].FillText22[0]":                        contact["zip"],
        "FW-003[0].Page2[0].PE_P2Header_gp[0].PersonWaivingName_ft[0]": contact["name"],
        "FW-003[0].Page2[0].PE_P2Header_gp[0].CaseNumber_ft[0]":        case.get("case_number", ""),
    }

    _write_pdf(
        template_path,
        output_path,
        values,
        field_appearances={
            "CaseName_ft[0]": "/Helv 7 Tf 0 g",
        },
    )
    _safe_print(f"  ✓ FW-003  → {output_path}")


# ─────────────────────────────────────────────────────────────
# SC-105  Proof of Service by Mail
# ─────────────────────────────────────────────────────────────

def fill_sc105(case, template_path, output_path):
    contact = _plaintiff_contact(case)
    d = case.get("defendant", DEFENDANT_DEFAULTS["city_of_oakland"])
    svc = case.get("service", {})
    generated_date = _forms_generated_date(case)

    values = {
        **_caption_fields(
            case,
            "SC-105[0].Page1[0].RightCaption[0].CourtInfo[0]",
            "SC-105[0].Page1[0].RightCaption[0].CaseNumber[0]",
            "SC-105[0].Page1[0].RightCaption[0].CaseName[0]",
        ),

        # Party names
        "SC-105[0].Page1[0].List1[0].Item[0].FullName3[0]": contact["name"],
        "SC-105[0].Page1[0].List1[0].Item[0].FullName2[0]": d.get("name", "City of Oakland"),

        # Signature
        "SC-105[0].Page1[0].Sign[0].SigDate4[0]": generated_date,
        "SC-105[0].Page1[0].Sign[0].SigName[0]":  svc.get("server_name", contact["name"]),

        **_caption_fields(
            case,
            "SC-105[0].Page2[0].RightCaption[0].CourtInfo[0]",
            "SC-105[0].Page2[0].RightCaption[0].CaseNumber[0]",
            "SC-105[0].Page2[0].RightCaption[0].CaseName[0]",
        ),
        "SC-105[0].Page2[0].List7[0].Item7[0].FullName10[0]":       contact["name"],
        "SC-105[0].Page2[0].List7[0].Item7[0].FullName12[0]":       d.get("name", "City of Oakland"),
    }

    _write_pdf(template_path, output_path, values)
    _safe_print(f"  ✓ SC-105  → {output_path}")


# ─────────────────────────────────────────────────────────────
# SC-109  Claim of Exemption / Request re: defendant
# ─────────────────────────────────────────────────────────────

def fill_sc109_exemption(case, template_path, output_path):
    contact = _plaintiff_contact(case)

    values = {
        **_caption_fields(
            case,
            "SC-109[0].Page1[0].Right_Caption[0].County[0].CourtInfo[0]",
            "SC-109[0].Page1[0].Right_Caption[0].CN[0].CaseNumber[0]",
            "SC-109[0].Page1[0].Right_Caption[0].CN[0].CaseName[0]",
        ),

        # Section 1: declarant
        "SC-109[0].Page1[0].List1[0].li1[0].NameField[0]":    contact["name"],
        "SC-109[0].Page1[0].List1[0].li1[0].Address[0]":      contact["street"],
        "SC-109[0].Page1[0].List1[0].li1[0].RelateField[0]":  "Plaintiff",

        # Section 2: check plaintiff
        "SC-109[0].Page1[0].List2[0].li1[0].PltfCheck[0]": "/Yes",
        "SC-109[0].Page1[0].List2[0].li1[0].PltfName[0]":  contact["name"],

        **_caption_fields(
            case,
            None,
            "SC-109[0].Page2[0].Header[0].CaseNumber[0]",
            "SC-109[0].Page2[0].Header[0].CaseName[0]",
        ),
    }

    _write_pdf(template_path, output_path, values)
    _safe_print(f"  ✓ SC-109 exemption → {output_path}")


# ─────────────────────────────────────────────────────────────
# SC-112A  Attachment to Plaintiff's Claim (itemized damages)
# ─────────────────────────────────────────────────────────────

def fill_sc112a(case, template_path, output_path):
    """Fill SC-112A, Proof of Service by Mail (Small Claims).

    SC-112A is completed and signed by a SERVER — an adult who is NOT a
    party to the case (Cal. Rules of Court, rule 3.2107) — so the server's
    information (item 1), the document checkboxes (item 2), the mailing
    date, and the signature are intentionally left blank for the server to
    complete. We prefill only what is known in advance: the case number
    and the names/addresses of the parties to be served (item 3b), i.e.
    the defendant and any additional defendants.
    """
    d = case.get("defendant", DEFENDANT_DEFAULTS["city_of_oakland"])

    def _mail_addr(dd):
        street = dd.get("address") or dd.get("street") or ""
        line2 = " ".join(x for x in [
            dd.get("city", ""), dd.get("state", ""), dd.get("zip", ""),
        ] if x)
        return ", ".join(part for part in [street, line2] if part)

    values = {
        "SC-112A[0].Page1[0].Header[0].CaseNumber_ft[0]": case.get("case_number", ""),
    }

    # Item 3b table — "Name of party served" / "Mailing address on the
    # envelope". Row field-name suffixes on the official form, top to bottom:
    _row_suffixes = ["11", "12", "13", "14", "1"]
    parties = [d] + list(case.get("additional_defendants") or [])
    for i, dd in enumerate(parties[:len(_row_suffixes)]):
        sfx = _row_suffixes[i]
        values[f"SC-112A[0].Page1[0].List3[0].Lib[0].Table[0].FillText10\\.{sfx}[0]"] = dd.get("name", "")
        values[f"SC-112A[0].Page1[0].List3[0].Lib[0].Table[0].FillText11\\.{sfx}[0]"] = _mail_addr(dd)

    _write_pdf(template_path, output_path, values)
    _safe_print(f"  ✓ SC-112A → {output_path}")


# ─────────────────────────────────────────────────────────────
# SC-150  Request to Postpone Trial (Small Claims)
# Code Civ. Proc. § 116.570; Cal. Rules of Court, rule 3.2107
# ─────────────────────────────────────────────────────────────

def fill_sc150(case, template_path, output_path):
    """Fill SC-150, Request to Postpone Trial, from case["postponement"].

    Expected keys in case["postponement"] (all optional except reason /
    requested_date, which make the request meaningful):
        requester_name      ignored; plaintiff name is always used at top
        role                "plaintiff" | "defendant" (default "plaintiff")
        mailing_address     ignored; plaintiff address is always used at top
        phone               ignored; plaintiff phone is always used at top
        current_trial_date  item 2 — date trial is now scheduled (MM/DD/YYYY)
        requested_date      item 3 — approximate new date requested
        reason              item 4 — why the postponement is needed
        late_reason         item 5 — why not requested sooner (trial < 10 days out)
        service_status      item 6 — "not_filed" (6a) | "served" (6b) |
                            "not_served" (6c) | "unknown" (6d)
        served              list of up to 2 {name, county, date} dicts (6b)
        unserved_names      list of up to 2 names (6c)
        unknown_names       list of up to 2 names (6d)
        request_date        signature date, defaults to forms generated date
    """
    contact = _plaintiff_contact(case)
    post = case.get("postponement", {}) or {}
    generated_date = _forms_generated_date(case)
    cn = _case_name(case)

    requester = contact["name"]
    role = (post.get("role") or "plaintiff").strip().lower()
    mailing = contact["mailing_address"]

    served   = [s for s in (post.get("served") or []) if s.get("name")][:2]
    unserved = [n for n in (post.get("unserved_names") or []) if n][:2]
    unknown  = [n for n in (post.get("unknown_names") or []) if n][:2]
    status = post.get("service_status") or ("served" if served else "")

    def srv(i, key):
        return served[i].get(key, "") if i < len(served) else ""

    def nth(seq, i):
        return seq[i] if i < len(seq) else ""

    values = {
        **_caption_fields(
            case,
            "SC-150[0].Page1[0].Caption_sf[0].supcourt[0].CourtInfo[0]",
            "SC-150[0].Page1[0].Caption_sf[0].casenumbername[0].CaseNumber[0]",
            "SC-150[0].Page1[0].Caption_sf[0].casenumbername[0].CaseName[0]",
        ),

        # 1. My name is / mailing address / phone / plaintiff-or-defendant
        "SC-150[0].Page1[0].List1[0].item1[0].FillText01[0]": requester,
        "SC-150[0].Page1[0].List1[0].item1[0].FillText03[0]": mailing,
        "SC-150[0].Page1[0].List1[0].item1[0].FillText04[0]": contact["phone"],
        "SC-150[0].Page1[0].List1[0].item1[0].CheckBox01[0]": "/1" if role == "plaintiff" else "/Off",
        "SC-150[0].Page1[0].List1[0].item1[0].CheckBox01[1]": "/2" if role == "defendant" else "/Off",

        # 2. My trial is now scheduled for (date)
        "SC-150[0].Page1[0].List2[0].item2[0].FillText05[0]": post.get("current_trial_date", ""),

        # 3. I ask the court to postpone my trial until (approximate date)
        "SC-150[0].Page1[0].List3[0].item3[0].FillText06[0]": post.get("requested_date", ""),

        # 4. I am asking for this postponement because (explain)
        "SC-150[0].Page1[0].List4[0].item4[0].FillText08[0]": post.get("reason", ""),

        # 5. If trial is within 10 days, why the request wasn't made sooner
        "SC-150[0].Page1[0].List5[0].item5[0].FillText15[0]": post.get("late_reason", ""),

        # 6. Has your claim been served?
        # 6a. No — I am a defendant and have not filed a claim
        "SC-150[0].Page1[0].List6[0].Lia[0].CheckBox04[0]": "/1" if status == "not_filed" else "/Off",
        # 6b. Yes — the parties listed below have been served
        "SC-150[0].Page1[0].List6[0].Lib[0].CheckBox04[0]": "/2" if status == "served" else "/Off",
        "SC-150[0].Page1[0].List6[0].Lib[0].sublistb[0].Li1[0].FillText1[0]":  srv(0, "name"),
        "SC-150[0].Page1[0].List6[0].Lib[0].sublistb[0].Li1[0].FillText2[0]":  srv(0, "county"),
        "SC-150[0].Page1[0].List6[0].Lib[0].sublistb[0].Li1[0].FillText3[0]":  srv(0, "date"),
        "SC-150[0].Page1[0].List6[0].Lib[0].sublistb[0].Li2[0].FillText16[0]": srv(1, "name"),
        "SC-150[0].Page1[0].List6[0].Lib[0].sublistb[0].Li2[0].FillText17[0]": srv(1, "county"),
        "SC-150[0].Page1[0].List6[0].Lib[0].sublistb[0].Li2[0].FillText18[0]": srv(1, "date"),
        # 6c. No — the parties listed below have not been served
        "SC-150[0].Page1[0].List6[0].Lic[0].CheckBox10[0]": "/3" if status == "not_served" else "/Off",
        "SC-150[0].Page1[0].List6[0].Lic[0].FillText19[0]": nth(unserved, 0),
        "SC-150[0].Page1[0].List6[0].Lic[0].FillText20[0]": nth(unserved, 1),
        # 6d. I do not know — clerk mailed the claim
        "SC-150[0].Page1[0].List6[0].Lid[0].CheckBox11[0]": "/4" if status == "unknown" else "/Off",
        "SC-150[0].Page1[0].List6[0].Lid[0].FillText21[0]": nth(unknown, 0),
        "SC-150[0].Page1[0].List6[0].Lid[0].FillText22[0]": nth(unknown, 1),

        # Signature (declaration under penalty of perjury)
        "SC-150[0].Page1[0].sign[0].Date1[0]":     generated_date,
        "SC-150[0].Page1[0].sign[0].printname[0]": requester,
    }

    _write_pdf(template_path, output_path, values)
    _safe_print(f"  ✓ SC-150  → {output_path}")


def has_postponement(case: dict) -> bool:
    """True if the case has enough postponement data to generate an SC-150."""
    post = case.get("postponement", {}) or {}
    return bool((post.get("reason") or "").strip() or (post.get("requested_date") or "").strip())


# ─────────────────────────────────────────────────────────────
# SC-107  Small Claims Subpoena and Declaration
# ─────────────────────────────────────────────────────────────

_SC107_DEFAULT_GOOD_CAUSE = (
    "The records requested are in the exclusive possession and control of the "
    "subpoenaed agency and cannot be obtained by any other means. The records "
    "are necessary to prepare for trial because they document the sweep at "
    "issue, the officers and employees who carried it out, the property that "
    "was seized or destroyed, and the policies under which the sweep was "
    "conducted."
)

_SC107_DEFAULT_MATERIALITY = (
    "This case concerns the seizure and destruction of plaintiff's personal "
    "property during an encampment sweep. The records requested show what "
    "happened during the sweep, who ordered and carried it out, what property "
    "was taken, whether the defendant followed its own policies and the "
    "requirements of due process, and the value and disposition of the "
    "property. Each category of records bears directly on the defendant's "
    "liability and on plaintiff's damages."
)


def _render_sc107_attachment_pages(case, output_path):
    """Render Attachment 2a, 3, and 4 pages for the SC-107 via reportlab."""
    from reportlab.lib.pagesizes import letter
    from reportlab.pdfgen.canvas import Canvas
    from reportlab.lib.units import inch
    from reportlab.lib.utils import simpleSplit

    sub = case.get("subpoena", {}) or {}
    requests = [r for r in sub.get("requests", []) if r]
    good_cause = (sub.get("good_cause") or "").strip() or _SC107_DEFAULT_GOOD_CAUSE
    materiality = (sub.get("materiality") or "").strip() or _SC107_DEFAULT_MATERIALITY

    c = Canvas(output_path, pagesize=letter)
    width, height = letter
    margin = inch * 0.75
    body_w = width - 2 * margin

    def header(label):
        y = height - margin
        c.setFont("Helvetica-Bold", 12)
        c.drawString(margin, y, f"ATTACHMENT {label}")
        c.setFont("Helvetica", 10)
        y -= 15
        c.drawString(margin, y, "to Form SC-107, Declaration in Support of Small Claims Subpoena")
        y -= 14
        c.drawString(margin, y, f"Case name: {_case_name(case)}")
        y -= 14
        c.drawString(margin, y, f"Case number: {case.get('case_number', '') or '____________________'}")
        y -= 26
        return y

    def attachment(label, blocks):
        y = header(label)
        c.setFont("Helvetica", 11)
        for block in blocks:
            for line in simpleSplit(block, "Helvetica", 11, body_w):
                if y < margin:
                    c.showPage()
                    y = header(f"{label} (continued)")
                    c.setFont("Helvetica", 11)
                c.drawString(margin, y, line)
                y -= 14
            y -= 7
        c.showPage()

    attachment("2a", [
        "Documents and other things to be produced by the witness (SC-107, item 2a):",
    ] + ([f"{i}. {r}" for i, r in enumerate(requests, 1)] or ["(No documents specified.)"]))

    attachment("3", [
        "Good cause exists for the production of the documents and other things "
        "described in item 2 for the following reasons (SC-107, item 3):",
        good_cause,
    ])

    attachment("4", [
        "These documents are material to the issues involved in this case for "
        "the following reasons (SC-107, item 4):",
        materiality,
    ])

    c.save()


def fill_sc107(case, template_path, output_path):
    """Prepare the SC-107 subpoena package: checkboxes only + attachments.

    The form's text fields are intentionally left blank. Only the "Continued
    on Attachment" checkboxes for items 2a, 3, and 4 (plus the item 2a "For
    trial or hearing" box) are checked, and the substantive content — the
    document requests, good cause, and materiality — is rendered on separate
    Attachment 2a / 3 / 4 pages appended behind the form.
    """
    values = {
        # Item 2a: "For trial or hearing" + "Continued on Attachment 2a."
        "SC-107[0].Page2[0].List2[0].Lia[0].CB\\.2\\.3\\.1\\.0[0]": "/Yes",
        "SC-107[0].Page2[0].List2[0].Lia[0].CB\\.2\\.3\\.1\\.1[0]": "/Yes",
        # Item 3: "Continued on Attachment 3."
        "SC-107[0].Page2[0].List3[0].item3[0].CB\\.2\\.30[0]": "/Yes",
        # Item 4: "Continued on Attachment 4."
        "SC-107[0].Page2[0].List4[0].item4[0].CB\\.222[0]": "/Yes",
    }

    _write_pdf(template_path, output_path, values)

    # Render the attachment pages and append them behind the form.
    att_path = output_path + ".attachments.pdf"
    _render_sc107_attachment_pages(case, att_path)
    main_bytes = Path(output_path).read_bytes()
    att_bytes = Path(att_path).read_bytes()
    merged = PdfWriter()

    # Keep backing streams alive until write completes. pypdf can keep
    # references to source streams while writing, and ephemeral BytesIO
    # instances may trigger "I/O operation on closed file".
    main_stream = io.BytesIO(main_bytes)
    att_stream = io.BytesIO(att_bytes)
    merged.append(PdfReader(main_stream))
    merged.append(PdfReader(att_stream))
    with open(output_path, "wb") as f:
        merged.write(f)
    main_stream.close()
    att_stream.close()
    try:
        os.remove(att_path)
    except OSError:
        pass
    _safe_print(f"  ✓ SC-107  → {output_path} (boxes checked; Attachments 2a, 3, 4 appended)")


# ─────────────────────────────────────────────────────────────
# SC-109  Authorization to Appear (Small Claims)
# ─────────────────────────────────────────────────────────────

def fill_sc109(case, template_path, output_path):
    """Fill SC-109 asking the court to let a helper assist the plaintiff.

    Uses `case['assistant']` (name, address, relationship, reason). Per the
    form's own instructions, item 3 is skipped and item 4 — the request to
    assist a plaintiff or defendant who cannot properly present their claim
    or defense (Code Civ. Proc. § 116.540) — is checked and explained.
    """
    p = case["plaintiff"]
    helper = case.get("assistant", {}) or {}
    generated_date = _forms_generated_date(case)

    values = {
        **_caption_fields(
            case,
            "SC-109[0].Page1[0].Right_Caption[0].County[0].CourtInfo[0]",
            "SC-109[0].Page1[0].Right_Caption[0].CN[0].CaseNumber[0]",
            "SC-109[0].Page1[0].Right_Caption[0].CN[0].CaseName[0]",
        ),
        **_caption_fields(
            case,
            None,
            "SC-109[0].Page2[0].Header[0].CaseNumber[0]",
            "SC-109[0].Page2[0].Header[0].CaseName[0]",
        ),

        # 1. Helper's name, address, and relationship to the plaintiff
        "SC-109[0].Page1[0].List1[0].li1[0].NameField[0]": helper.get("name", ""),
        "SC-109[0].Page1[0].List1[0].li1[0].Address[0]": helper.get("address", ""),
        "SC-109[0].Page1[0].List1[0].li1[0].RelateField[0]": helper.get("relationship", ""),

        # 2. Appearing for the plaintiff
        "SC-109[0].Page1[0].List2[0].li1[0].PltfCheck[0]": "/1",
        "SC-109[0].Page1[0].List2[0].li1[0].PltfName[0]": p.get("name", ""),

        # 4. Request to assist a party who cannot properly present their
        #    claim or defense (item 3 intentionally left blank)
        "SC-109[0].Page2[0].List4[0].li1[0].Ch4[0]": "/1",
        "SC-109[0].Page2[0].List4[0].li1[0].Field12[0]": helper.get("reason", ""),

        # 5. Date + printed name (the helper signs)
        "SC-109[0].Page2[0].List5[0].li1[0].FillText8[0]": generated_date,
        "SC-109[0].Page2[0].List5[0].li1[0].FillText9[0]": helper.get("name", ""),
    }

    _write_pdf(template_path, output_path, values)
    _safe_print(f"  ✓ SC-109  → {output_path}")


# ─────────────────────────────────────────────────────────────
# SC-100A  Other Plaintiffs or Defendants (generated if no template)
# ─────────────────────────────────────────────────────────────

def _render_sc100a_reportlab(party: dict, case: dict, output_path: str, role: str = "defendant"):
    """Create a simple SC-100A-looking PDF using reportlab.

    Fields filled: case number, party name, street, city/state/zip, mailing
    address, phone, job title, checkbox for fictitious name, and signature line.
    """
    try:
        from reportlab.lib.pagesizes import letter
        from reportlab.pdfgen.canvas import Canvas
        from reportlab.lib.units import inch
    except Exception:
        # If reportlab isn't available, write a plain text fallback PDF via pypdf
        from pypdf import PdfWriter
        w = PdfWriter()
        w.add_blank_page(width=612, height=792)
        os.makedirs(os.path.dirname(output_path) or '.', exist_ok=True)
        with open(output_path, 'wb') as f:
            w.write(f)
        return

    c = Canvas(output_path, pagesize=letter)
    width, height = letter
    generated_date = _forms_generated_date(case)

    def tx(x, y, text, size=10):
        c.setFont("Helvetica", size)
        c.drawString(x, y, text)

    y = height - inch * 0.75
    tx(inch * 0.5, y, "SC-100A  Other Plaintiffs or Defendants", size=14)
    y -= 18
    tx(inch * 0.5, y, f"Attached to SC-100 case number: {case.get('case_number','')}")
    y -= 24

    # Section for this extra party
    tx(inch * 0.5, y, f"Role: {role.title()}")
    y -= 16
    tx(inch * 0.5, y, f"Name: {party.get('name','')}")
    y -= 14
    tx(inch * 0.5, y, f"Street address: {party.get('street','')}")
    y -= 14
    city = party.get('city','')
    state = party.get('state','')
    zipc = party.get('zip','')
    tx(inch * 0.5, y, f"City: {city}    State: {state}    Zip: {zipc}")
    y -= 14
    tx(inch * 0.5, y, f"Mailing address (if different): {party.get('mailing','')}")
    y -= 14
    tx(inch * 0.5, y, f"Phone: {party.get('phone','')}")
    y -= 14
    tx(inch * 0.5, y, f"Job title, if known: {party.get('job_title','')}")
    y -= 20

    # Fictitious name checkbox line
    tx(inch * 0.5, y, "Is this plaintiff/defendant doing business under a fictitious name?  [ ] Yes   [ ] No")
    y -= 30

    tx(inch * 0.5, y, "I declare under penalty of perjury under California state law that the information above and on any attachments is true and correct.")
    y -= 28
    tx(inch * 0.5, y, f"Date: {generated_date}     Type or print your name: ________________________     Sign your name: ________________________")

    c.showPage()
    c.save()


def fill_sc100a_for_party(case, output_path, party: dict, role: str = "defendant"):
    """Public wrapper to create SC-100A PDF for a single party.

    If an official `templates/sc100a.pdf` exists in the templates dir, we could
    prefer to use it; currently we render a simple matching layout via reportlab.
    """
    _render_sc100a_reportlab(party, case, output_path, role=role)
    _safe_print(f"  ✓ SC-100A → {output_path}")


# ─────────────────────────────────────────────────────────────
# TEMPLATE & META PATHS
# ─────────────────────────────────────────────────────────────

TEMPLATES = {
    "sc100":  "templates/sc100.pdf",
    "sc105":  "templates/sc105.pdf",
    "sc109":  "templates/sc109.pdf",
    "sc112a": "templates/sc112a.pdf",
    "sc150":  "templates/sc150.pdf",
    "fw001":  "templates/fw001.pdf",
    "fw003":  "templates/fw003.pdf",
    "sc107":  "templates/sc107.pdf",
}

FIELD_META = {
    "sc100": "field_meta/sc100_fields.json",
    "fw001": "field_meta/fw001_fields.json",
}


# ─────────────────────────────────────────────────────────────
# CASE RUNNER
# ─────────────────────────────────────────────────────────────

def fill_case(case_path):
    with open(case_path) as f:
        case = json.load(f)

    try:
        validate_case(case)
    except ValueError as e:
        _safe_print(f"\n[SKIP] {case_path}\n{e}")
        return False

    name_slug = re.sub(r"[^a-z0-9]+", "_", case["plaintiff"]["name"].lower()).strip("_")
    _safe_print(f"\n→ {case['plaintiff']['name']}")

    fill_sc100(case,  TEMPLATES["sc100"],  f"output/{name_slug}_sc100.pdf",  FIELD_META["sc100"])
    fill_fw001(case,  TEMPLATES["fw001"],  f"output/{name_slug}_fw001.pdf",  FIELD_META["fw001"])
    fill_fw003(case,  TEMPLATES["fw003"],  f"output/{name_slug}_fw003.pdf")
    fill_sc112a(case, TEMPLATES["sc112a"], f"output/{name_slug}_sc112a.pdf")

    # SC-150 Request to Postpone Trial — only if case has postponement data
    if has_postponement(case):
        fill_sc150(case, TEMPLATES["sc150"], f"output/{name_slug}_sc150.pdf")

    # Proof of Service — only if case has service data
    if case.get("service", {}).get("service_date"):
        fill_sc105(case, TEMPLATES["sc105"], f"output/{name_slug}_sc105.pdf")

    # SC-109 — only if explicitly requested
    if case.get("default_request"):
        fill_sc109(case, TEMPLATES["sc109"], f"output/{name_slug}_sc109.pdf")

    return True


# ─────────────────────────────────────────────────────────────
# CASE TEMPLATE GENERATOR
# ─────────────────────────────────────────────────────────────

CASE_TEMPLATE = {
    "_comment": "Oakland encampment property destruction — edit TODO fields, keep Oakland defaults",

    "plaintiff": {
        "name":   "TODO: Full Legal Name",
        "street": "c/o TODO: shelter/address (use c/o for unhoused clients)",
        "city":   "Oakland",
        "state":  "CA",
        "zip":    "TODO: 5-digit ZIP",
        "phone":  "TODO: 510-XXX-XXXX",
        "email":  "",
    },

    "defendant": {
        "name":          "City of Oakland",
        "address":       "One Frank H. Ogawa Plaza",
        "city":          "Oakland",
        "state":         "CA",
        "zip":           "94612",
        "agent_name":    "City Clerk",
        "agent_title":   "City Clerk",
        "agent_address": "One Frank H. Ogawa Plaza",
        "agent_city":    "Oakland",
        "agent_state":   "CA",
        "agent_zip":     "94612",
    },

    "claim": {
        "amount":               "TODO: Dollar amount up to 12500",
        "reason":               "TODO: Describe the sweep — when, where, what was taken/destroyed, why it was wrongful",
        "incident_date":        "TODO: MM/DD/YYYY",
        "date_started":         "",
        "date_through":         "",
        "damages_calculation":  "TODO: Itemize property value + emotional distress damages",
        "govt_claim_filed_date": "TODO: MM/DD/YYYY — date government tort claim filed with City Clerk",
        "items": [
            {"description": "TODO: e.g. Tent and sleeping bag", "value": "TODO: 350"},
            {"description": "TODO: e.g. Clothing (jacket, shoes, 3 sets)", "value": "TODO: 500"},
            {"description": "TODO: e.g. Personal documents and ID", "value": "TODO: 200"},
        ],
    },

    "filing": {
        "filing_date":      "TODO: MM/DD/YYYY",
        "demanded_payment": True,
    },

    "fee_waiver": {
        "basis":                 "5c",
        "waive_option":          "all",
        "receives_medi_cal":     False,
        "receives_snap":         False,
        "receives_calworks":     False,
        "income_source_1":       "TODO: e.g. General Assistance, SSI, none",
        "income_amount_1":       "TODO: monthly amount",
        "total_monthly_income":  "TODO: total",
        "expense_housing":       "0",
        "expense_food":          "TODO",
        "expense_utilities":     "0",
        "expense_medical":       "TODO",
        "expense_transport":     "TODO",
        "total_monthly_expenses": "TODO",
    },

    "declaration": {
        "declarant_name": "TODO: Full Legal Name",
        "content": (
            "TODO: I am the plaintiff in this action. On [date], the City of Oakland "
            "Department of Public Works conducted an encampment sweep at [location]. "
            "Without adequate notice or opportunity to retrieve my belongings, City employees "
            "seized and destroyed my personal property including [list items]. "
            "I declare under penalty of perjury under the laws of the State of California "
            "that the foregoing is true and correct."
        ),
    },

    "postponement": {
        "_comment": "Optional — fill only to generate SC-150 Request to Postpone Trial",
        "role":               "plaintiff",
        "current_trial_date": "",
        "requested_date":     "",
        "reason":             "",
        "late_reason":        "",
        "service_status":     "",
        "served":             [],
        "unserved_names":     [],
        "unknown_names":      [],
        "request_date":       "",
    },
}


def generate_template(output_path):
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    with open(output_path, "w") as f:
        json.dump(CASE_TEMPLATE, f, indent=2)
    _safe_print(f"Template written → {output_path}")
    _safe_print(f"Edit the TODO fields, then run:  python fill_forms.py {output_path}")


# ─────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Fill Oakland small claims forms from a case JSON file.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "  python fill_forms.py cases/jane_doe.json\n"
            "  python fill_forms.py cases/\n"
            "  python fill_forms.py --new cases/new_client.json"
        ),
    )
    parser.add_argument(
        "target",
        nargs="?",
        help="Case JSON file, or directory of JSON files.",
    )
    parser.add_argument(
        "--new",
        metavar="OUTPUT",
        help="Generate a blank case template at OUTPUT path.",
    )
    args = parser.parse_args()

    if args.new:
        generate_template(args.new)
        return

    if not args.target:
        parser.print_help()
        sys.exit(1)

    target = Path(args.target)
    if target.is_dir():
        cases = sorted(target.glob("*.json"))
        if not cases:
            _safe_print(f"No JSON files found in {target}")
            sys.exit(1)
        results = [fill_case(f) for f in cases]
        ok = sum(results)
        _safe_print(f"\nDone: {ok}/{len(cases)} cases filled successfully.")
        if ok < len(cases):
            sys.exit(1)
    else:
        if not fill_case(target):
            sys.exit(1)
        _safe_print("\nDone.")


if __name__ == "__main__":
    main()
