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
import sys
import os
import re
import argparse
from pathlib import Path
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, BooleanObject, DictionaryObject

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



# ─────────────────────────────────────────────────────────────
# FIELD METADATA (checkbox on/off values)
# ─────────────────────────────────────────────────────────────

def load_field_meta(json_path):
    with open(json_path) as f:
        fields = json.load(f)
    return {item["field_id"]: item for item in fields}


# ─────────────────────────────────────────────────────────────
# VALIDATION
# ─────────────────────────────────────────────────────────────

def validate_case(case):
    """Raise ValueError with clear message if required fields are missing."""
    errors = []

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
    if not claim.get("govt_claim_filed_date"):
        errors.append(
            "claim.govt_claim_filed_date is required when suing a public entity"
        )

    filing = case.get("filing", {})
    if not filing.get("filing_date"):
        errors.append("filing.filing_date is required")

    if errors:
        raise ValueError(
            "Case validation failed:\n" + "\n".join(f"  • {e}" for e in errors)
        )


# ─────────────────────────────────────────────────────────────
# CORE PDF WRITE
# ─────────────────────────────────────────────────────────────

def _write_pdf(template_path, output_path, values):
    reader = PdfReader(template_path)
    writer = PdfWriter()
    writer.append(reader)

    for page in writer.pages:
        writer.update_page_form_field_values(page, values)

    # Ensure PDF viewers render the updated field values by setting
    # the AcroForm /NeedAppearances flag. Some viewers require this
    # to generate appearance streams for filled fields when the PDF
    # is opened or saved elsewhere.
    try:
        root = writer._root_object
        # Set the AcroForm /NeedAppearances flag so PDF viewers render filled fields
        try:
            acro = root.get(NameObject("/AcroForm")) if hasattr(root, 'get') else root[NameObject("/AcroForm")] if NameObject("/AcroForm") in root else None
        except Exception:
            acro = None

        if acro is None:
            root.update({NameObject("/AcroForm"): DictionaryObject({NameObject("/NeedAppearances"): BooleanObject(True)})})
        else:
            try:
                acro.update({NameObject("/NeedAppearances"): BooleanObject(True)})
            except Exception:
                root.update({NameObject("/AcroForm"): DictionaryObject({NameObject("/NeedAppearances"): BooleanObject(True)})})
    except Exception:
        # Non-fatal: continue writing even if we can't set the flag
        pass

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    with open(output_path, "wb") as f:
        writer.write(f)

    # Try to flatten the PDF so appearances are visually baked in for all viewers.
    # Flattening is optional and will be skipped if PyMuPDF (`pymupdf`) isn't installed.
    try:
        _flatten_pdf(output_path)
    except Exception:
        pass


def _case_name(case):
    d = case.get("defendant", DEFENDANT_DEFAULTS["city_of_oakland"])
    return f"{case['plaintiff']['name']} v. {d.get('name', 'City of Oakland')}"


# ─────────────────────────────────────────────────────────────
# SC-100  Plaintiff's Claim and ORDER to Go to Small Claims Court
# ─────────────────────────────────────────────────────────────

def fill_sc100(case, template_path, output_path, field_meta_path):
    meta = load_field_meta(field_meta_path)
    p = case["plaintiff"]
    d = case.get("defendant", DEFENDANT_DEFAULTS["city_of_oakland"])
    claim = case["claim"]
    filing = case.get("filing", {})

    def cb(fid, checked):
        if checked:
            return meta.get(fid, {}).get("checked_value", "/Yes")
        return meta.get(fid, {}).get("unchecked_value", "/Off")

    values = {
        # Header
        "SC-100[0].Page1[0].CaptionRight[0].County[0].CourtInfo[0]": _court_info(case),

        # Plaintiff
        "SC-100[0].Page2[0].List1[0].Item1[0].PlaintiffName1[0]":    p["name"],
        "SC-100[0].Page2[0].List1[0].Item1[0].PlaintiffAddress1[0]": p.get("street", ""),
        "SC-100[0].Page2[0].List1[0].Item1[0].PlaintiffCity1[0]":    p.get("city", ""),
        "SC-100[0].Page2[0].List1[0].Item1[0].PlaintiffState1[0]":   p.get("state", "CA"),
        "SC-100[0].Page2[0].List1[0].Item1[0].PlaintiffZip1[0]":     p.get("zip", ""),
        "SC-100[0].Page2[0].List1[0].Item1[0].PlaintiffPhone1[0]":   p.get("phone", ""),
        "SC-100[0].Page2[0].List1[0].Item1[0].EmailAdd1[0]":         p.get("email", ""),

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
        "SC-100[0].Page3[0].List4[0].Item4[0].Checkbox50[0]": cb(
            "SC-100[0].Page3[0].List4[0].Item4[0].Checkbox50[0]",
            filing.get("demanded_payment", True),
        ),
        "SC-100[0].Page3[0].List4[0].Item4[0].Checkbox50[1]": cb(
            "SC-100[0].Page3[0].List4[0].Item4[0].Checkbox50[1]",
            not filing.get("demanded_payment", True),
        ),

        # Venue — where property was damaged
        "SC-100[0].Page3[0].List5[0].Lib[0].Checkbox5cb[0]": cb(
            "SC-100[0].Page3[0].List5[0].Lib[0].Checkbox5cb[0]", True
        ),
        "SC-100[0].Page3[0].List6[0].item6[0].ZipCode1[0]": _venue_zip(case),

        # Attorney-client fee dispute? No
        "SC-100[0].Page3[0].List7[0].item7[0].Checkbox60[0]": cb(
            "SC-100[0].Page3[0].List7[0].item7[0].Checkbox60[0]", False
        ),
        "SC-100[0].Page3[0].List7[0].item7[0].Checkbox60[1]": cb(
            "SC-100[0].Page3[0].List7[0].item7[0].Checkbox60[1]", True
        ),

        # Suing a public entity? Yes + claim date
        "SC-100[0].Page3[0].List8[0].item8[0].Checkbox61[0]": cb(
            "SC-100[0].Page3[0].List8[0].item8[0].Checkbox61[0]", True
        ),
        "SC-100[0].Page3[0].List8[0].item8[0].Checkbox61[1]": cb(
            "SC-100[0].Page3[0].List8[0].item8[0].Checkbox61[1]", False
        ),
        "SC-100[0].Page3[0].List8[0].item8[0].Date4[0]": claim.get("govt_claim_filed_date", ""),

        # Filed more than 12 claims this year? No
        "SC-100[0].Page4[0].List9[0].Item9[0].Checkbox62[0]": cb(
            "SC-100[0].Page4[0].List9[0].Item9[0].Checkbox62[0]", False
        ),
        "SC-100[0].Page4[0].List9[0].Item9[0].Checkbox62[1]": cb(
            "SC-100[0].Page4[0].List9[0].Item9[0].Checkbox62[1]", True
        ),

        # Claim for more than $2,500? Yes
        "SC-100[0].Page4[0].List10[0].li10[0].Checkbox63[0]": cb(
            "SC-100[0].Page4[0].List10[0].li10[0].Checkbox63[0]", True
        ),
        "SC-100[0].Page4[0].List10[0].li10[0].Checkbox63[1]": cb(
            "SC-100[0].Page4[0].List10[0].li10[0].Checkbox63[1]", False
        ),

        # Signature
        "SC-100[0].Page4[0].Sign[0].Date1[0]":          filing.get("filing_date", ""),
        "SC-100[0].Page4[0].Sign[0].PlaintiffName1[0]":  p["name"],

        # Repeated caption fields
        "SC-100[0].Page2[0].PxCaption[0].Plaintiff[0]": p["name"],
        "SC-100[0].Page3[0].PxCaption[0].Plaintiff[0]": p["name"],
        "SC-100[0].Page4[0].PxCaption[0].Plaintiff[0]": p["name"],
    }

    _write_pdf(template_path, output_path, values)
    try:
        _flatten_using_field_meta(output_path, field_meta_path)
    except Exception:
        pass
    print(f"  ✓ SC-100  → {output_path}")


# ─────────────────────────────────────────────────────────────
# FW-001  Request to Waive Court Fees
# ─────────────────────────────────────────────────────────────

def fill_fw001(case, template_path, output_path, field_meta_path):
    meta = load_field_meta(field_meta_path)
    p = case["plaintiff"]
    fw = case.get("fee_waiver", {})
    filing = case.get("filing", {})

    def cb(fid, checked):
        if checked:
            return meta.get(fid, {}).get("checked_value", "/Yes")
        return meta.get(fid, {}).get("unchecked_value", "/Off")

    basis = fw.get("basis", "5c")  # "5a" | "5b" | "5c"

    values = {
        # Header
        "FW-001[0].Page1[0].RightCaption[0].CourtInfo[0]": _court_info(case),
        "FW-001[0].Page1[0].RightCaption[0].CaseName[0]":  _case_name(case),

        # Section 1: Petitioner info
        "FW-001[0].Page1[0].List1[0].item1[0].PetitionerName1[0]":      p["name"],
        "FW-001[0].Page1[0].List1[0].item1[0].PetitionerStrAddress[0]": p.get("street", ""),
        "FW-001[0].Page1[0].List1[0].item1[0].PetitionerCity[0]":       p.get("city", ""),
        "FW-001[0].Page1[0].List1[0].item1[0].PetitionerState[0]":      p.get("state", "CA"),
        "FW-001[0].Page1[0].List1[0].item1[0].PetitionerZip[0]":        p.get("zip", ""),
        "FW-001[0].Page1[0].List1[0].item1[0].PetitionerTel[0]":        p.get("phone", ""),

        # Section 4: Superior Court fees
        "FW-001[0].Page1[0].List4[0].item4[0].WaiveSuperiorCrtFee[0]": cb(
            "FW-001[0].Page1[0].List4[0].item4[0].WaiveSuperiorCrtFee[0]", True
        ),

        # Section 5a: Public benefits
        "FW-001[0].Page1[0].List5[0].Lia[0].PublicBenefitReceived[0]": cb(
            "FW-001[0].Page1[0].List5[0].Lia[0].PublicBenefitReceived[0]", basis == "5a"
        ),
        "FW-001[0].Page1[0].List5[0].Lia[0].PublicBenefitSNAP[0]": cb(
            "FW-001[0].Page1[0].List5[0].Lia[0].PublicBenefitSNAP[0]",
            fw.get("receives_snap", False),
        ),
        "FW-001[0].Page1[0].List5[0].Lia[0].PublicBenefitMediCal[0]": cb(
            "FW-001[0].Page1[0].List5[0].Lia[0].PublicBenefitMediCal[0]",
            fw.get("receives_medi_cal", False),
        ),
        "FW-001[0].Page1[0].List5[0].Lia[0].PublicBenefitCalWORKSTANF[0]": cb(
            "FW-001[0].Page1[0].List5[0].Lia[0].PublicBenefitCalWORKSTANF[0]",
            fw.get("receives_calworks", False),
        ),

        # Section 5b: Gross income below threshold
        "FW-001[0].Page1[0].List5[0].Lib[0].GrossMonthIncomeLess[0]": cb(
            "FW-001[0].Page1[0].List5[0].Lib[0].GrossMonthIncomeLess[0]", basis == "5b"
        ),

        # Section 5c: Cannot afford fees (most common for our clients)
        "FW-001[0].Page1[0].List5[0].Lic[0].IncomeInsufficientRequest[0]": cb(
            "FW-001[0].Page1[0].List5[0].Lic[0].IncomeInsufficientRequest[0]", basis == "5c"
        ),
        "FW-001[0].Page1[0].List5[0].Lic[0].FeeRequestDef[0]": cb(
            "FW-001[0].Page1[0].List5[0].Lic[0].FeeRequestDef[0]",
            basis == "5c" and fw.get("waive_option", "all") == "all",
        ),

        # Signature
        "FW-001[0].Page1[0].Sign[0].SigDate[0]":        filing.get("filing_date", ""),
        "FW-001[0].Page1[0].Sign[0].PetitionerName[0]": p["name"],

        # Page 2 caption + income
        "FW-001[0].Page2[0].pXCaption[0].PetitionerName1[0]":           p["name"],
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
        _flatten_using_field_meta(output_path, field_meta_path)
    except Exception:
        pass
    print(f"  ✓ FW-001  → {output_path}")


# ─────────────────────────────────────────────────────────────
# FW-003  Order on Court Fee Waiver  (court-completed — we pre-fill header only)
# ─────────────────────────────────────────────────────────────

def fill_fw003(case, template_path, output_path):
    p = case["plaintiff"]

    values = {
        "FW-003[0].Page1[0].Stamp_court_case[0].CourtInfo_ft[0]":  _court_info(case),
        "FW-003[0].Page1[0].Stamp_court_case[0].CaseNumber_ft[0]": case.get("case_number", ""),
        "FW-003[0].Page1[0].Stamp_court_case[0].CaseName_ft[0]":   _case_name(case),
        "FW-003[0].Page1[0].PersonWaivingName_ft[0]":              p["name"],
        "FW-003[0].Page2[0].PE_P2Header_gp[0].PersonWaivingName_ft[0]": p["name"],
        "FW-003[0].Page2[0].PE_P2Header_gp[0].CaseNumber_ft[0]":        case.get("case_number", ""),
    }

    _write_pdf(template_path, output_path, values)
    print(f"  ✓ FW-003  → {output_path}")


# ─────────────────────────────────────────────────────────────
# SC-105  Proof of Service by Mail
# ─────────────────────────────────────────────────────────────

def fill_sc105(case, template_path, output_path):
    p = case["plaintiff"]
    d = case.get("defendant", DEFENDANT_DEFAULTS["city_of_oakland"])
    svc = case.get("service", {})
    cn = _case_name(case)

    values = {
        # Page 1 caption
        "SC-105[0].Page1[0].RightCaption[0].CourtInfo[0]":  _court_info(case),
        "SC-105[0].Page1[0].RightCaption[0].CaseNumber[0]": case.get("case_number", ""),
        "SC-105[0].Page1[0].RightCaption[0].CaseName[0]":   cn,

        # Party names
        "SC-105[0].Page1[0].List1[0].Item[0].FullName3[0]": p["name"],
        "SC-105[0].Page1[0].List1[0].Item[0].FullName2[0]": d.get("name", "City of Oakland"),

        # Signature
        "SC-105[0].Page1[0].Sign[0].SigDate4[0]": svc.get("service_date", ""),
        "SC-105[0].Page1[0].Sign[0].SigName[0]":  svc.get("server_name", p["name"]),

        # Page 2 caption + party names
        "SC-105[0].Page2[0].RightCaption[0].CourtInfo[0]":          _court_info(case),
        "SC-105[0].Page2[0].RightCaption[0].CaseNumber[0]":         case.get("case_number", ""),
        "SC-105[0].Page2[0].RightCaption[0].CaseName[0]":           cn,
        "SC-105[0].Page2[0].List7[0].Item7[0].FullName10[0]":       p["name"],
        "SC-105[0].Page2[0].List7[0].Item7[0].FullName12[0]":       d.get("name", "City of Oakland"),
    }

    _write_pdf(template_path, output_path, values)
    print(f"  ✓ SC-105  → {output_path}")


# ─────────────────────────────────────────────────────────────
# SC-109  Claim of Exemption / Request re: defendant
# ─────────────────────────────────────────────────────────────

def fill_sc109(case, template_path, output_path):
    p = case["plaintiff"]
    cn = _case_name(case)

    values = {
        # Caption
        "SC-109[0].Page1[0].Right_Caption[0].County[0].CourtInfo[0]": _court_info(case),
        "SC-109[0].Page1[0].Right_Caption[0].CN[0].CaseNumber[0]":    case.get("case_number", ""),
        "SC-109[0].Page1[0].Right_Caption[0].CN[0].CaseName[0]":      cn,

        # Section 1: declarant
        "SC-109[0].Page1[0].List1[0].li1[0].NameField[0]":    p["name"],
        "SC-109[0].Page1[0].List1[0].li1[0].Address[0]":      p.get("street", ""),
        "SC-109[0].Page1[0].List1[0].li1[0].RelateField[0]":  "Plaintiff",

        # Section 2: check plaintiff
        "SC-109[0].Page1[0].List2[0].li1[0].PltfCheck[0]": "/Yes",
        "SC-109[0].Page1[0].List2[0].li1[0].PltfName[0]":  p["name"],

        # Page 2 header
        "SC-109[0].Page2[0].Header[0].CaseName[0]":   cn,
        "SC-109[0].Page2[0].Header[0].CaseNumber[0]": case.get("case_number", ""),
    }

    _write_pdf(template_path, output_path, values)
    print(f"  ✓ SC-109  → {output_path}")


# ─────────────────────────────────────────────────────────────
# SC-112A  Attachment to Plaintiff's Claim (itemized damages)
# ─────────────────────────────────────────────────────────────

def fill_sc112a(case, template_path, output_path):
    p = case["plaintiff"]
    claim = case["claim"]
    filing = case.get("filing", {})
    items = claim.get("items", [])

    def _item_desc(i):
        if i < len(items):
            it = items[i]
            desc = it.get("description", "")
            val = it.get("value", "")
            return f"{desc}  ${val}" if val else desc
        return ""

    values = {
        "SC-112A[0].Page1[0].Header[0].CaseNumber_ft[0]": case.get("case_number", ""),

        # Item 1 block: plaintiff name, case name, incident date, then item rows
        "SC-112A[0].Page1[0].List1[0].Item1[0].FillText01[0]": p["name"],
        "SC-112A[0].Page1[0].List1[0].Item1[0].FillText02[0]": _case_name(case),
        "SC-112A[0].Page1[0].List1[0].Item1[0].FillText03[0]": claim.get("incident_date", ""),
        "SC-112A[0].Page1[0].List1[0].Item1[0].FillText04[0]": _item_desc(0),
        "SC-112A[0].Page1[0].List1[0].Item1[0].FillText05[0]": _item_desc(1),
        "SC-112A[0].Page1[0].List1[0].Item1[0].FillText06[0]": _item_desc(2),
        "SC-112A[0].Page1[0].List1[0].Item1[0].FillText07[0]": _item_desc(3),
        "SC-112A[0].Page1[0].List1[0].Item1[0].FillText08[0]": _item_desc(4),

        # Signature
        "SC-112A[0].Page1[0].Sign[0].FillText14[0]": filing.get("filing_date", ""),
        "SC-112A[0].Page1[0].Sign[0].FillText16[0]": p["name"],
    }

    _write_pdf(template_path, output_path, values)
    print(f"  ✓ SC-112A → {output_path}")


# ─────────────────────────────────────────────────────────────
# SC-150  Declaration
# ─────────────────────────────────────────────────────────────

def fill_sc150(case, template_path, output_path):
    p = case["plaintiff"]
    claim = case["claim"]
    filing = case.get("filing", {})
    decl = case.get("declaration", {})
    cn = _case_name(case)

    declarant = decl.get("declarant_name", p["name"])
    content = decl.get("content", claim.get("reason", ""))

    values = {
        # Caption
        "SC-150[0].Page1[0].Caption_sf[0].supcourt[0].CourtInfo[0]":              _court_info(case),
        "SC-150[0].Page1[0].Caption_sf[0].casenumbername[0].CaseNumber[0]":       case.get("case_number", ""),
        "SC-150[0].Page1[0].Caption_sf[0].casenumbername[0].CaseName[0]":         cn,

        # Section 1: declarant identifies themselves
        "SC-150[0].Page1[0].List1[0].item1[0].FillText01[0]": declarant,
        "SC-150[0].Page1[0].List1[0].item1[0].FillText03[0]": claim.get("incident_date", ""),
        "SC-150[0].Page1[0].List1[0].item1[0].FillText04[0]": content,

        # Additional paragraphs (optional)
        "SC-150[0].Page1[0].List2[0].item2[0].FillText05[0]": decl.get("paragraph_2", ""),
        "SC-150[0].Page1[0].List3[0].item3[0].FillText06[0]": decl.get("paragraph_3", ""),
        "SC-150[0].Page1[0].List4[0].item4[0].FillText08[0]": decl.get("paragraph_4", ""),
        "SC-150[0].Page1[0].List5[0].item5[0].FillText15[0]": decl.get("paragraph_5", ""),

        # Signature
        "SC-150[0].Page1[0].sign[0].Date1[0]":      filing.get("filing_date", ""),
        "SC-150[0].Page1[0].sign[0].printname[0]":  declarant,
    }

    _write_pdf(template_path, output_path, values)
    print(f"  ✓ SC-150  → {output_path}")


# ─────────────────────────────────────────────────────────────
# SC-107  Small Claims Subpoena and Declaration
# ─────────────────────────────────────────────────────────────

def fill_sc107(case, template_path, output_path):
    """Fill SC-107 subpoena using `case['subpoena']` fields.

    We populate the case caption, party names, recipient/custodian/service
    address, the requested production list (joined into a text area), and
    signature/date fields. The SC-107 template contains many widget names;
    we set the most useful editable fields so the form is usable.
    """
    p = case["plaintiff"]
    d = case.get("defendant", DEFENDANT_DEFAULTS["city_of_oakland"])
    sub = case.get("subpoena", {}) or {}

    requests = [r for r in sub.get("requests", []) if r]
    requests_text = "\n".join(requests)

    values = {
        # Header / caption: party names + case number
        "SC-107[0].Page2[0].Header[0].TitlePartyName[0].Party1_ft[0]": p.get("name", ""),
        "SC-107[0].Page2[0].Header[0].TitlePartyName[0].Party2_ft[0]": d.get("name", ""),
        "SC-107[0].Page2[0].Header[0].CaseNumber[0].CaseNumber_ft[0]": case.get("case_number", ""),

        # Caption lines on page 1
        "SC-107[0].Page1[0].Caption_sf[0].Plaintiff[0].Gr1[0].T813[0]": p.get("name", ""),
        "SC-107[0].Page1[0].Caption_sf[0].Defendant[0].Gr3[0].T9[0]": d.get("name", ""),

        # Recipient / custodian / service location
        "SC-107[0].Page1[0].List1[0].Lib[0].Field3[0]": sub.get("to", ""),
        "SC-107[0].Page1[0].List2[0].Field4[0]": sub.get("custodian", ""),
        "SC-107[0].Page1[0].List2[0].Field5[0]": sub.get("service_location", ""),

        # Put the requests into a large free-text area on page 3
        "SC-107[0].Page3[0].List3[0].Lih[0].TextField51[0]": requests_text,

        # Signature block (date + printed name)
        "SC-107[0].Page2[0].Sign[0].Date1[0]": case.get("filing", {}).get("filing_date", ""),
        "SC-107[0].Page2[0].Sign[0].PrintName[0]": p.get("name", ""),
    }

    _write_pdf(template_path, output_path, values)
    print(f"  ✓ SC-107  → {output_path}")


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
    tx(inch * 0.5, y, "Date: ________________________     Type or print your name: ________________________     Sign your name: ________________________")

    c.showPage()
    c.save()


def fill_sc100a_for_party(case, output_path, party: dict, role: str = "defendant"):
    """Public wrapper to create SC-100A PDF for a single party.

    If an official `templates/sc100a.pdf` exists in the templates dir, we could
    prefer to use it; currently we render a simple matching layout via reportlab.
    """
    _render_sc100a_reportlab(party, case, output_path, role=role)
    print(f"  ✓ SC-100A → {output_path}")


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
        print(f"\n[SKIP] {case_path}\n{e}")
        return False

    name_slug = re.sub(r"[^a-z0-9]+", "_", case["plaintiff"]["name"].lower()).strip("_")
    print(f"\n→ {case['plaintiff']['name']}")

    fill_sc100(case,  TEMPLATES["sc100"],  f"output/{name_slug}_sc100.pdf",  FIELD_META["sc100"])
    fill_fw001(case,  TEMPLATES["fw001"],  f"output/{name_slug}_fw001.pdf",  FIELD_META["fw001"])
    fill_fw003(case,  TEMPLATES["fw003"],  f"output/{name_slug}_fw003.pdf")
    fill_sc112a(case, TEMPLATES["sc112a"], f"output/{name_slug}_sc112a.pdf")
    fill_sc150(case,  TEMPLATES["sc150"],  f"output/{name_slug}_sc150.pdf")

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
}


def generate_template(output_path):
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    with open(output_path, "w") as f:
        json.dump(CASE_TEMPLATE, f, indent=2)
    print(f"Template written → {output_path}")
    print(f"Edit the TODO fields, then run:  python fill_forms.py {output_path}")


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
            print(f"No JSON files found in {target}")
            sys.exit(1)
        results = [fill_case(f) for f in cases]
        ok = sum(results)
        print(f"\nDone: {ok}/{len(cases)} cases filled successfully.")
        if ok < len(cases):
            sys.exit(1)
    else:
        if not fill_case(target):
            sys.exit(1)
        print("\nDone.")


if __name__ == "__main__":
    main()
