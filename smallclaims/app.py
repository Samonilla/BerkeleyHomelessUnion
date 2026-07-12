"""
Oakland Encampment Small Claims — Streamlit UI

Run:  streamlit run app.py
"""

import contextlib
import io
import json
import os
import re
import sys
import tempfile
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st
from dateutil import parser as _dateutil
from docx import Document

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))

from fill_forms import (
    fill_sc100, fill_fw001, fill_fw003, fill_sc112a, fill_sc150,
    fill_sc105, fill_sc107, fill_sc100a_for_party, validate_case, DEFENDANT_DEFAULTS,
)
from courts import ALL_COUNTIES, courthouses_for_county, court_info_string
from defendants import ALL_CITIES, defendant_info

_META_SC100 = str(HERE / "field_meta" / "sc100_fields.json")
_META_FW001 = str(HERE / "field_meta" / "fw001_fields.json")
_TPL = HERE / "templates"


# ─── PDF helpers ──────────────────────────────────────────────────────────────

def _slug(name: str) -> str:
    return re.sub(r"[^a-z0-9]+", "_", name.lower()).strip("_") or "case"


def _normalize_plain_language(text: str) -> str:
    text = re.sub(r"\s+", " ", text or "").strip()
    if not text:
        return ""
    text = text[0].upper() + text[1:]
    if text[-1] not in ".!?":
        text += "."
    return text


def _build_guided_declaration(text: str, answers: dict) -> str:
    intro = [
        "I am the plaintiff in this action.",
        "I am submitting this declaration under penalty of perjury and state that the following is true and correct.",
    ]
    facts = []

    if answers.get("incident_date"):
        facts.append(f"On {answers['incident_date']}, the events described below occurred.")

    for key, value in answers.items():
        if key in {"incident_date", "items", "claim_amount"}:
            continue
        cleaned = _normalize_plain_language(value)
        if cleaned:
            facts.append(cleaned)

    items = answers.get("items") or []
    if items:
        item_lines = []
        for item in items:
            description = str(item.get("description") or "").strip()
            value = str(item.get("value") or "").strip()
            condition = str(item.get("condition") or "").strip() or "Unknown"
            if description and value:
                item_lines.append(
                    f"I lost {description}, which was valued at ${value}, and it was in {condition.lower()} condition when it was destroyed."
                )
            elif description:
                item_lines.append(
                    f"I lost {description}, and it was in {condition.lower()} condition when it was destroyed."
                )
        if item_lines:
            facts.append(" ".join(item_lines))

    if text.strip():
        facts.append(_normalize_plain_language(text))

    if answers.get("claim_amount"):
        facts.append(f"The total value of the property I lost is approximately ${answers['claim_amount']}.")

    paragraphs = intro + [f"{i}. {fact}" for i, fact in enumerate(facts, start=1)]
    paragraphs.append("I declare under penalty of perjury that the foregoing is true and correct.")
    return "\n\n".join(paragraphs)


def _build_declaration_docx(text: str) -> bytes:
    document = Document()
    document.add_heading("Declaration", level=1)
    for paragraph_text in text.split("\n\n"):
        if paragraph_text.strip():
            document.add_paragraph(paragraph_text)

    buf = io.BytesIO()
    document.save(buf)
    return buf.getvalue()


def _build_govt_claim_docx(data: dict) -> bytes:
    """Build a Gov. Code §§ 905/910 claim-for-damages form as a Word document."""
    entity = (data.get("entity") or "").strip()
    document = Document()
    document.add_heading("CLAIM FOR DAMAGES AGAINST A PUBLIC ENTITY", level=1)
    to_text = f"To: Office of the City Clerk{', ' + entity if entity else ''}"
    if data.get("clerk_address"):
        to_text += f"\n{data['clerk_address']}"
    document.add_paragraph(to_text)
    document.add_paragraph(
        "This claim is presented pursuant to California Government Code "
        "sections 905, 910, and 910.2."
    )

    def _row(label, value):
        p = document.add_paragraph()
        p.add_run(f"{label}: ").bold = True
        p.add_run(value if value else "____________________________________________")

    _row("1. Claimant name", data.get("claimant_name"))
    _row("2. Mailing address (send all notices here)", data.get("claimant_address"))
    _row("3. Phone", data.get("claimant_phone"))
    _row("4. Date of occurrence", data.get("incident_date"))
    _row("5. Place of occurrence", data.get("incident_location"))
    _row("6. Circumstances of the occurrence", data.get("description"))
    _row("7. General description of the injury, damage, or loss", data.get("description"))
    _row("8. Names of public employees or agencies causing the loss, if known",
         data.get("employees"))

    amount_raw = (data.get("amount") or "").replace("$", "").replace(",", "").strip()
    try:
        amount_val = float(amount_raw)
    except ValueError:
        amount_val = None
    if amount_val is not None and amount_val > 10000:
        _row(
            "9. Amount claimed",
            "The amount claimed exceeds $10,000 and this would be a limited "
            "civil case. (Gov. Code § 910(f).)",
        )
    else:
        _row(
            "9. Amount claimed as of presentation, with basis of computation",
            (f"${amount_raw} — the value of the claimant's personal property "
             "destroyed or taken, as described above.") if amount_raw else "",
        )

    document.add_paragraph("")
    document.add_paragraph(
        "I declare under penalty of perjury under the laws of the State of "
        "California that the foregoing is true and correct."
    )
    document.add_paragraph("Dated: ____________________")
    document.add_paragraph(
        f"Signature: ____________________    {data.get('claimant_name') or ''}"
    )

    buf = io.BytesIO()
    document.save(buf)
    return buf.getvalue()


_DEFAULT_SUBPOENA_REQUESTS = [
    "All body-worn camera, dashboard camera, and other video or audio recordings from the sweep.",
    "All incident reports, field notes, supplemental reports, and emails by officers involved.",
    "All dispatch logs, radio transmissions, and 911/311 call recordings for the incident.",
    "All complaints, investigations, internal affairs files, and disciplinary records for involved officers.",
    "All policies, training materials, written directives, and encampment sweep protocols.",
    "All property seizure, storage, chain-of-custody, and disposal records.",
    "All internal communications, memos, and coordination records between police, DPW, and other City agencies.",
    "All surveillance camera and private video footage from the sweep location.",
    "All records authorizing, scheduling, or directing the sweep.",
]


@contextlib.contextmanager
def _quiet():
    with open(os.devnull, "w") as nul:
        old, sys.stdout = sys.stdout, nul
        try:
            yield
        finally:
            sys.stdout = old


def _generate_pdfs(case: dict) -> dict:
    """Fill all forms. Returns {label: bytes}. Raises ValueError on bad input."""
    validate_case(case)
    result = {}
    with tempfile.TemporaryDirectory() as td:
        tmp = Path(td)
        with _quiet():
            fill_sc100(case, str(_TPL/"sc100.pdf"), str(tmp/"sc100.pdf"), _META_SC100)
            result["SC-100"] = (tmp/"sc100.pdf").read_bytes()

            fill_fw001(case, str(_TPL/"fw001.pdf"), str(tmp/"fw001.pdf"), _META_FW001)
            result["FW-001"] = (tmp/"fw001.pdf").read_bytes()

            fill_fw003(case, str(_TPL/"fw003.pdf"), str(tmp/"fw003.pdf"))
            result["FW-003"] = (tmp/"fw003.pdf").read_bytes()

            fill_sc112a(case, str(_TPL/"sc112a.pdf"), str(tmp/"sc112a.pdf"))
            result["SC-112A"] = (tmp/"sc112a.pdf").read_bytes()

            # SC-150 form generation
            try:
                fill_sc150(case, str(_TPL/"sc150.pdf"), str(tmp/"sc150.pdf"))
                result["SC-150"] = (tmp/"sc150.pdf").read_bytes()
            except Exception:
                # Handle SC-150 generation failure
                pass

            # SC-107 subpoena (if subpoena info present)
            try:
                fill_sc107(case, str(_TPL/"sc107.pdf"), str(tmp/"sc107.pdf"))
                result["SC-107"] = (tmp/"sc107.pdf").read_bytes()
            except Exception:
                # Non-fatal: continue generating other forms even if SC-107 fails
                pass

            # SC-100A: generate one form per additional defendant (if present)
            for i, ad in enumerate(case.get('additional_defendants', []) or [] , start=1):
                try:
                    outp = tmp/f"sc100a_defendant_{i}.pdf"
                    fill_sc100a_for_party(case, str(outp), ad, role='defendant')
                    result[f"SC-100A-DEF-{i}"] = outp.read_bytes()
                except Exception:
                    pass
    return result


def _make_zip(pdfs: dict, slug: str, flatten: bool = False) -> bytes:
    """Create a ZIP of the provided PDF bytes.

    By default (`flatten=False`) the function packages the exact in-memory
    PDF bytes (so the ZIP matches the individual PDF downloads shown in-UI).
    If `flatten=True` the function will try to rasterize + sanitize PDFs for
    maximum viewer compatibility.
    """
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for label, data in pdfs.items():
            out_bytes = data

            if flatten:
                # Try to rasterize/flatten PDF bytes so appearances are baked in for
                # viewers that ignore /NeedAppearances. If PyMuPDF isn't available,
                # fall back to the original bytes.
                try:
                    import fitz
                    # Open original bytes, render each page to an image-PDF, then
                    # combine pages into a new PDF bytes object.
                    doc = fitz.open(stream=data, filetype="pdf")
                    new = fitz.open()
                    mat = fitz.Matrix(2.0, 2.0)
                    for page in doc:
                        pix = page.get_pixmap(matrix=mat)
                        img_pdf = fitz.open("pdf", pix.tobytes("pdf"))
                        new.insert_pdf(img_pdf)
                    try:
                        out_bytes = new.write()
                    except Exception:
                        # Fallback: save to temp file then read
                        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmpf:
                            new.save(tmpf.name)
                            tmpname = tmpf.name
                        with open(tmpname, "rb") as tf:
                            out_bytes = tf.read()
                        os.unlink(tmpname)
                    new.close()
                    doc.close()
                except Exception:
                    out_bytes = data

                # Post-process to remove any remaining AcroForm and page /Annots
                # so ZIPs contain image-only PDFs that render reliably in viewers.
                try:
                    from pypdf import PdfReader, PdfWriter
                    from pypdf.generic import NameObject
                    rdr = PdfReader(io.BytesIO(out_bytes))
                    w = PdfWriter()
                    w.append(rdr)

                    # Remove page annotations
                    for p in w.pages:
                        try:
                            if '/Annots' in p:
                                p.pop('/Annots', None)
                        except Exception:
                            pass

                    # Remove AcroForm from root if present
                    try:
                        root = w._root_object
                        if NameObject('/AcroForm') in root:
                            try:
                                del root[NameObject('/AcroForm')]
                            except Exception:
                                try:
                                    root.pop(NameObject('/AcroForm'), None)
                                except Exception:
                                    pass
                    except Exception:
                        pass

                    bio = io.BytesIO()
                    w.write(bio)
                    out_bytes = bio.getvalue()
                except Exception:
                    # If post-processing fails, keep the original bytes
                    pass

            zf.writestr(f"{slug}_{label.lower().replace('-', '')}.pdf", out_bytes)
    return buf.getvalue()


def _show_downloads(pdfs: dict, slug: str, label: str = "") -> None:
    prefix = f"{label} — " if label else ""
    st.success(f"{prefix}Generated {len(pdfs)} forms.")
    st.download_button(
        "⬇️  Download All Forms (ZIP)",
        data=_make_zip(pdfs, slug, flatten=False),
        file_name=f"{slug}_forms.zip",
        mime="application/zip",
        type="primary",
        width="stretch",
        key=f"zip_{slug}",
    )
    # Optional: flattened ZIP for viewers that ignore form appearances
    st.download_button(
        "⬇️  Download All Forms (ZIP, flattened for compatibility)",
        data=_make_zip(pdfs, slug, flatten=True),
        file_name=f"{slug}_forms_flattened.zip",
        mime="application/zip",
        type="secondary",
        width="stretch",
        key=f"zip_flat_{slug}",
    )
    cols = st.columns(len(pdfs))
    for col, (lbl, data) in zip(cols, pdfs.items()):
        fname = f"{slug}_{lbl.lower().replace('-', '')}.pdf"
        with col:
            st.download_button(
                lbl, data=data, file_name=fname,
                mime="application/pdf", width="stretch",
                key=f"pdf_{slug}_{lbl}",
            )


# ─── Address / date parsing ───────────────────────────────────────────────────

_STATE_ZIP_RE  = re.compile(r"\b([A-Z]{2})\s+(\d{5}(?:-\d{4})?)\b")
_ZIP_RE        = re.compile(r"\b(\d{5})\b")
_KNOWN_CITIES  = re.compile(
    r"\b(Oakland|Emeryville|Richmond|Alameda|Hayward|Fremont|Berkeley|Newark|"
    r"Martinez|San Leandro|Walnut Creek|Pleasanton|Livermore|Castro Valley)\b",
    re.IGNORECASE,
)


def _parse_address(raw: str) -> dict:
    """Split a freeform US address into {street, city, state, zip}."""
    raw = str(raw).strip() if raw and not (isinstance(raw, float) and pd.isna(raw)) else ""
    if not raw:
        return {"street": "", "city": "Oakland", "state": "CA", "zip": ""}

    m = _STATE_ZIP_RE.search(raw)
    if m:
        state, zip_ = m.group(1), m.group(2)
        before = raw[: m.start()].rstrip(", ").strip()

        # If there's a comma, it separates street from city cleanly
        if "," in before:
            parts = [p.strip() for p in before.rsplit(",", 1)]
            street, city = parts[0], parts[1]
        else:
            # Find the last known city name in the text before the state
            city_match = None
            for cm in _KNOWN_CITIES.finditer(before):
                city_match = cm
            if city_match:
                street = before[: city_match.start()].strip()
                city = city_match.group(0).title()
            else:
                # Fallback: last word is city
                parts = before.rsplit(None, 1)
                street, city = (parts[0], parts[1]) if len(parts) == 2 else (before, "Oakland")

        return {"street": street or raw, "city": city, "state": state, "zip": zip_}

    # No STATE ZIP found — just extract ZIP if present
    z = _ZIP_RE.search(raw)
    return {"street": raw, "city": "Oakland", "state": "CA", "zip": z.group(1) if z else ""}


def _parse_date(raw) -> str:
    """Parse messy date strings → 'MM/DD/YYYY'. Returns '' on failure."""
    if not raw or (isinstance(raw, float) and pd.isna(raw)):
        return ""
    s = str(raw).strip()
    if not s:
        return ""
    # Handle ranges like "June 3 and 4 2025" → "June 3 2025"
    s = re.sub(r"\s+and\s+\d+", "", s)
    # Remove ordinal suffixes
    s = re.sub(r"(\d+)(st|nd|rd|th)\b", r"\1", s)
    # Remove slashes in "June 17/2026"
    s = re.sub(r"(\d+)/(\d{4})", r"\1 \2", s)
    try:
        return _dateutil.parse(s).strftime("%m/%d/%Y")
    except Exception:
        return s


# ─── Spreadsheet format detection ────────────────────────────────────────────

_INTAKE_COLS = {"Name", "Address", "Phone Number", "Location of Injury", "Date of Injury"}
_TEMPLATE_COL = "plaintiff_name"


def _detect_format(df: pd.DataFrame) -> str:
    cols = set(df.columns)
    if _INTAKE_COLS.issubset(cols):
        return "oakland_intake"
    if _TEMPLATE_COL in cols:
        return "template"
    return "unknown"


# ─── Oakland intake → case dict ───────────────────────────────────────────────

_CLAIM_REASON_TMPL = (
    "On {date}, the City of Oakland Department of Public Works conducted an encampment "
    "sweep at {location}, Oakland, CA. City employees seized and destroyed Plaintiff's "
    "personal property without adequate notice, without providing an opportunity to retrieve "
    "belongings, and without following the City's bag-and-tag policy. Plaintiff's property "
    "included clothing, sleeping equipment, tools, and personal documents. The City's actions "
    "violated the Fourth Amendment, Article I § 13 of the California Constitution, and the "
    "City's own policies."
)

_DECL_TMPL = (
    "I am the plaintiff in this action. On {date}, the City of Oakland Department of Public "
    "Works conducted an encampment sweep at {location}, Oakland, CA. Without adequate notice "
    "and without providing me an opportunity to retrieve my belongings, City employees seized "
    "and destroyed my personal property. My property included clothing, sleeping equipment, "
    "tools, and personal documents. At no time did City employees tag or store my property "
    "for later retrieval as required by the City's bag-and-tag policy. I have been unable to "
    "replace most of these items and have suffered ongoing hardship as a result. "
    "I declare under penalty of perjury under the laws of the State of California that the "
    "foregoing is true and correct."
)


def intake_row_to_case(row: pd.Series, defaults: dict) -> dict:
    def g(col, fallback=""):
        v = row.get(col, fallback)
        return fallback if (isinstance(v, float) and pd.isna(v)) else str(v).strip()

    addr = _parse_address(g("Address"))
    inc_date = _parse_date(g("Date of Injury"))
    location = g("Location of Injury")

    reason = defaults.get("claim_reason") or _CLAIM_REASON_TMPL.format(
        date=inc_date or "the date of the sweep",
        location=location or "the encampment",
    )
    decl = defaults.get("declaration") or _DECL_TMPL.format(
        date=inc_date or "the date of the sweep",
        location=location or "the encampment",
    )
    damages = defaults.get("damages_calculation") or (
        f"Property destroyed at encampment sweep at {location}. "
        f"Total estimated damages: ${defaults.get('claim_amount', '10000')}."
    )

    # Phone: take first number if multiple listed
    phone_raw = g("Phone Number")
    phone = re.split(r"[;,/]", phone_raw)[0].strip()

    return {
        "court": defaults.get("court", {}),
        "plaintiff": {
            "name":   g("Name"),
            "street": addr["street"],
            "city":   addr["city"],
            "state":  addr["state"],
            "zip":    addr["zip"],
            "phone":  phone,
            "email":  g("email"),
        },
        "defendant": defaults.get("defendant") or DEFENDANT_DEFAULTS["city_of_oakland"],
        "claim": {
            "amount":                defaults.get("claim_amount", "10000"),
            "reason":                reason,
            "incident_date":         inc_date,
            "damages_calculation":   damages,
            "govt_claim_filed_date": defaults.get("govt_claim_filed_date", ""),
            "items":                 [],
        },
        "filing": {
            "filing_date":      defaults.get("filing_date", ""),
            "demanded_payment": True,
        },
        "fee_waiver": {
            "basis":                  defaults.get("fw_basis", "5c"),
            "waive_option":           "all",
            "receives_medi_cal":      defaults.get("receives_medi_cal", False),
            "receives_snap":          defaults.get("receives_snap", False),
            "receives_calworks":      defaults.get("receives_calworks", False),
            "income_source_1":        defaults.get("income_source", ""),
            "income_amount_1":        defaults.get("income_amount", ""),
            "total_monthly_income":   defaults.get("total_income", ""),
            "expense_housing":        defaults.get("expense_housing", "0"),
            "expense_food":           defaults.get("expense_food", "0"),
            "expense_utilities":      "0",
            "expense_medical":        defaults.get("expense_medical", "0"),
            "expense_transport":      defaults.get("expense_transport", "0"),
            "total_monthly_expenses": defaults.get("total_expenses", ""),
        },
        "declaration": {
            "declarant_name": g("Name"),
            "content":        decl,
        },
    }


# ─── Generic template format → case dict ─────────────────────────────────────

def template_row_to_case(row: pd.Series) -> dict:
    def s(col, default=""):
        v = row.get(col, default)
        return default if (isinstance(v, float) and pd.isna(v)) else str(v).strip()

    def b(col, default=False):
        v = row.get(col, "")
        if isinstance(v, float) and pd.isna(v):
            return default
        return str(v).strip().upper() in ("TRUE", "YES", "1", "Y")

    items = []
    for i in range(1, 7):
        desc = s(f"item_{i}_desc")
        val  = s(f"item_{i}_value")
        if desc:
            items.append({"description": desc, "value": val})

    reason = s("claim_reason")
    return {
        "plaintiff": {
            "name":   s("plaintiff_name"),
            "street": s("plaintiff_street"),
            "city":   s("plaintiff_city", "Oakland"),
            "state":  s("plaintiff_state", "CA"),
            "zip":    s("plaintiff_zip"),
            "phone":  s("plaintiff_phone"),
            "email":  s("plaintiff_email"),
        },
        "defendant": DEFENDANT_DEFAULTS["city_of_oakland"],  # template format always uses spreadsheet data
        "claim": {
            "amount":                s("claim_amount"),
            "reason":                reason,
            "incident_date":         s("incident_date"),
            "damages_calculation":   s("damages_calculation") or reason,
            "govt_claim_filed_date": s("govt_claim_filed_date"),
            "items":                 items,
        },
        "filing": {
            "filing_date":      s("filing_date"),
            "demanded_payment": b("demanded_payment", True),
        },
        "fee_waiver": {
            "basis":                  s("fee_waiver_basis", "5c"),
            "waive_option":           "all",
            "receives_medi_cal":      b("receives_medi_cal"),
            "receives_snap":          b("receives_snap"),
            "receives_calworks":      b("receives_calworks"),
            "income_source_1":        s("income_source_1"),
            "income_amount_1":        s("income_amount_1"),
            "total_monthly_income":   s("total_monthly_income"),
            "expense_housing":        s("expense_housing", "0"),
            "expense_food":           s("expense_food", "0"),
            "expense_utilities":      "0",
            "expense_medical":        s("expense_medical", "0"),
            "expense_transport":      s("expense_transport", "0"),
            "total_monthly_expenses": s("total_monthly_expenses"),
        },
        "declaration": {
            "declarant_name": s("plaintiff_name"),
            "content":        s("declaration_content") or reason,
        },
        "subpoena": {
            "case_caption":          s("subpoena_case_caption"),
            "to":                    s("subpoena_to"),
            "custodian":             s("subpoena_custodian"),
            "service_location":      s("subpoena_service_location"),
            "requests": [
                s("subpoena_request_1"),
                s("subpoena_request_2"),
                s("subpoena_request_3"),
                s("subpoena_request_4"),
                s("subpoena_request_5"),
                s("subpoena_request_6"),
                s("subpoena_request_7"),
                s("subpoena_request_8"),
                s("subpoena_request_9"),
                s("subpoena_request_10"),
            ],
        },
    }


# ─── CSV template for download ────────────────────────────────────────────────

_TEMPLATE_COLS = [
    ("plaintiff_name",          "Full legal name",                                True,  "Jane Doe"),
    ("plaintiff_street",        "Street / PO Box (c/o ... for unhoused)",         True,  "c/o 1234 Telegraph Ave"),
    ("incident_date",           "Date of sweep MM/DD/YYYY",                       True,  "05/12/2025"),
    ("claim_amount",            "Total claim dollars (max 12500)",                True,  "10000"),
    ("claim_reason",            "What happened — used on SC-100 and SC-150",      True,  "On May 12 2025 City of Oakland DPW..."),
    ("govt_claim_filed_date",   "Date govt tort claim filed with City Clerk",     True,  "08/15/2025"),
    ("filing_date",             "Date filing court papers MM/DD/YYYY",            True,  "09/15/2025"),
    ("total_monthly_income",    "Total monthly income $",                          True,  "400"),
    ("total_monthly_expenses",  "Total monthly expenses $",                        True,  "300"),
    ("plaintiff_city",          "City (default Oakland)",                         False, "Oakland"),
    ("plaintiff_state",         "State (default CA)",                             False, "CA"),
    ("plaintiff_zip",           "ZIP code",                                       False, "94609"),
    ("plaintiff_phone",         "Phone number",                                   False, "510-555-0100"),
    ("plaintiff_email",         "Email",                                          False, ""),
    ("damages_calculation",     "How damages were calculated (SC-100)",         False, "Clothing $500..."),
    ("income_source_1",         "Primary income source",                          False, "General Assistance"),
    ("income_amount_1",         "Primary income amount $",                         False, "400"),
    ("expense_food",            "Monthly food/supplies $",                         False, "200"),
    ("expense_medical",         "Monthly medical $",                               False, "50"),
    ("expense_transport",       "Monthly transport $",                             False, "50"),
    ("expense_housing",         "Monthly housing $",                               False, "0"),
    ("receives_medi_cal",       "Receives Medi-Cal? TRUE/FALSE",                  False, "TRUE"),
    ("fee_waiver_basis",        "Fee waiver basis: 5a 5b or 5c",                  False, "5c"),
    ("declaration_content",     "First-person declaration (optional)",            False, ""),

    # SC-107 subpoena helper fields
    ("subpoena_case_caption",   "SC-107 subpoena case caption",                   False, "Jane Doe v. City of Oakland"),
    ("subpoena_to",             "Subpoena recipient / agency",                    False, "Oakland Police Department"),
    ("subpoena_custodian",      "Custodian of records",                            False, "Records Division"),
    ("subpoena_service_location","Service address for subpoena",                   False, "1515 Clay St, Oakland CA"),
    ("subpoena_request_1",      "SC-107 request item 1",                          False, "All body-worn camera and officer dashboard footage from the sweep."),
    ("subpoena_request_2",      "SC-107 request item 2",                          False, "All police incident reports, notes, and supplemental reports related to the sweep."),
    ("subpoena_request_3",      "SC-107 request item 3",                          False, "All dispatch logs, radio transmissions, and 911/311 call recordings for the incident."),
    ("subpoena_request_4",      "SC-107 request item 4",                          False, "All complaints, investigations, and disciplinary records for involved officers."),
    ("subpoena_request_5",      "SC-107 request item 5",                          False, "All internal communications, emails, memos, and directives regarding encampment sweeps."),
    ("subpoena_request_6",      "SC-107 request item 6",                          False, "All policies, training materials, use-of-force guidelines, and homeless encampment protocols."),
    ("subpoena_request_7",      "SC-107 request item 7",                          False, "All property seizure, storage, chain-of-custody, and disposal records."),
    ("subpoena_request_8",      "SC-107 request item 8",                          False, "All surveillance camera and private video footage from the sweep location."),
    ("subpoena_request_9",      "SC-107 request item 9",                          False, "All records of coordination between police, DPW, and other City agencies."),
    ("subpoena_request_10",     "SC-107 request item 10",                         False, "All logs, schedules, and written directives authorizing the sweeps."),
    ("item_1_desc",             "Property item 1 description",                    False, "Tent and sleeping bag"),
    ("item_1_value",            "Property item 1 value $",                         False, "350"),
    ("item_2_desc",             "Property item 2 description",                    False, "Clothing"),
    ("item_2_value",            "Property item 2 value $",                         False, "500"),
]


def _csv_template_bytes() -> bytes:
    row = {c[0]: c[3] for c in _TEMPLATE_COLS}
    buf = io.StringIO()
    pd.DataFrame([row]).to_csv(buf, index=False)
    return buf.getvalue().encode()


_SC107_TEMPLATE_COLS = [
    ("subpoena_case_caption",    "SC-107 subpoena case caption",                   False, "Jane Doe v. City of Oakland"),
    ("subpoena_to",              "Subpoena recipient / agency",                    False, "Oakland Police Department"),
    ("subpoena_custodian",       "Custodian of records",                            False, "Records Division"),
    ("subpoena_service_location","Service address for subpoena",                   False, "1515 Clay St, Oakland CA"),
    ("subpoena_request_1",       "SC-107 request item 1",                          False, "All body-worn camera and officer dashboard footage from the sweep."),
    ("subpoena_request_2",       "SC-107 request item 2",                          False, "All police incident reports, notes, and supplemental reports related to the sweep."),
    ("subpoena_request_3",       "SC-107 request item 3",                          False, "All dispatch logs, radio transmissions, and 911/311 call recordings for the incident."),
    ("subpoena_request_4",       "SC-107 request item 4",                          False, "All complaints, investigations, and disciplinary records for involved officers."),
    ("subpoena_request_5",       "SC-107 request item 5",                          False, "All internal communications, emails, memos, and directives regarding encampment sweeps."),
    ("subpoena_request_6",       "SC-107 request item 6",                          False, "All policies, training materials, use-of-force guidelines, and homeless encampment protocols."),
    ("subpoena_request_7",       "SC-107 request item 7",                          False, "All property seizure, storage, chain-of-custody, and disposal records."),
    ("subpoena_request_8",       "SC-107 request item 8",                          False, "All surveillance camera and private video footage from the sweep location."),
    ("subpoena_request_9",       "SC-107 request item 9",                          False, "All records of coordination between police, DPW, and other City agencies."),
    ("subpoena_request_10",      "SC-107 request item 10",                         False, "All logs, schedules, and written directives authorizing the sweeps."),
]


def _sc107_csv_template_bytes() -> bytes:
    row = {c[0]: c[3] for c in _SC107_TEMPLATE_COLS}
    buf = io.StringIO()
    pd.DataFrame([row]).to_csv(buf, index=False)
    return buf.getvalue().encode()


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE
# ═══════════════════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="CA Small Claims Autofiller",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.title("California Encampment — Small Claims Autofiller")
st.caption(
    "Generates a government claim form plus SC-100, SC-100A, FW-001, FW-003, "
    "SC-112A, SC-150, and SC-107 for encampment property destruction cases in "
    "any California county — from the initial government claim, to filing the "
    "lawsuit, to preparing for trial."
)


# ─── Court selector widget (reused in both tabs) ─────────────────────────────

def _court_selector(key_prefix: str, default_county: str = "Alameda") -> dict:
    """Render county + courthouse dropdowns. Returns a court dict for case["court"]."""
    col_county, col_house = st.columns([1, 2])
    with col_county:
        county = st.selectbox(
            "County *",
            ALL_COUNTIES,
            index=ALL_COUNTIES.index(default_county) if default_county in ALL_COUNTIES else 0,
            key=f"{key_prefix}_county",
        )
    houses = courthouses_for_county(county)
    house_labels = [f"{h['city']} — {h['name']}" for h in houses]
    with col_house:
        house_idx = st.selectbox(
            "Courthouse *",
            range(len(house_labels)),
            format_func=lambda i: house_labels[i],
            key=f"{key_prefix}_house",
        )
    chosen = houses[house_idx]
    return {
        "county":  county,
        "name":    chosen["name"],
        "address": chosen["address"],
        "city":    chosen["city"],
        "zip":     chosen["zip"],
    }

_SELECTOR_MANUAL = "🏘️ Unincorporated area / Other (enter manually)"


def _defendant_selector(key_prefix: str, default_city: str = "Oakland") -> dict:
    """Dropdown of all CA municipalities + manual entry for unincorporated areas.
    Always returns a fully-populated defendant dict.
    """
    city_options = [_SELECTOR_MANUAL] + ALL_CITIES
    default_idx = ALL_CITIES.index(default_city) + 1 if default_city in ALL_CITIES else 1
    selected = st.selectbox(
        "Defendant City / Municipality *",
        city_options,
        index=default_idx,
        key=f"{key_prefix}_def_city",
        help=(
            "Select the California city or town being sued. "
            "For unincorporated county areas, county agencies, or any other entity, "
            "choose 'Unincorporated area / Other' to enter all fields manually."
        ),
    )
    if selected == _SELECTOR_MANUAL:
        st.caption("Enter the defendant's information manually.")
        c1, c2 = st.columns(2)
        with c1:
            name_v  = st.text_input("Defendant Name *",    placeholder="County of Alameda", key=f"{key_prefix}_cust_name")
            str_v   = st.text_input("Street Address",       placeholder="1221 Oak St",       key=f"{key_prefix}_cust_street")
            city_v  = st.text_input("City",                 placeholder="Oakland",           key=f"{key_prefix}_cust_city")
            s1, s2  = st.columns(2)
            with s1:
                state_v = st.text_input("State", value="CA",          key=f"{key_prefix}_cust_state")
            with s2:
                zip_v   = st.text_input("ZIP",   placeholder="94612", key=f"{key_prefix}_cust_zip")
        with c2:
            agent_name_v   = st.text_input("Agent for Service",  value="Clerk",         key=f"{key_prefix}_cust_agent_name")
            agent_title_v  = st.text_input("Agent Title",        value="County Clerk",  key=f"{key_prefix}_cust_agent_title")
            agent_str_v    = st.text_input("Agent Street",       placeholder="1221 Oak St", key=f"{key_prefix}_cust_agent_street")
            agent_city_v   = st.text_input("Agent City",         placeholder="Oakland",     key=f"{key_prefix}_cust_agent_city")
            agent_zip_v    = st.text_input("Agent ZIP",          placeholder="94612",       key=f"{key_prefix}_cust_agent_zip")
        return {
            "name":          name_v.strip(),
            "address":       str_v.strip(),
            "city":          city_v.strip(),
            "state":         state_v.strip() or "CA",
            "zip":           zip_v.strip(),
            "agent_name":    agent_name_v.strip() or "Clerk",
            "agent_title":   agent_title_v.strip() or "County Clerk",
            "agent_address": agent_str_v.strip(),
            "agent_city":    agent_city_v.strip(),
            "agent_state":   "CA",
            "agent_zip":     agent_zip_v.strip(),
        }
    d = defendant_info(selected)
    st.caption(
        f"**{d['name']}** · {d['address']}, {d['city']}, CA {d['zip']} "
        f"· Agent for service: {d['agent_name']}"
    )
    return d


_CUSTOM_DEFENDANT = "🏘️ Unincorporated area / Other (enter manually)"


def _defendant_block(key_prefix: str, def_id: int, is_primary: bool) -> dict:
    """One defendant entry: city dropdown (or enter-your-own) plus editable
    address fields prefilled from the municipality database.

    The primary defendant (SC-100) also gets agent-for-service fields;
    additional defendants (SC-100A) get phone / mailing / job-title fields.
    """
    city_options = [_CUSTOM_DEFENDANT] + ALL_CITIES
    default_idx = (ALL_CITIES.index("Oakland") + 1) if (is_primary and "Oakland" in ALL_CITIES) else 0
    selected = st.selectbox(
        "Defendant City / Municipality *",
        city_options,
        index=default_idx,
        key=f"{key_prefix}_def{def_id}_city_sel",
        help=(
            "Select the California city or town being sued, or choose "
            "'Unincorporated area / Other' for residents of unincorporated county areas, "
            "county agencies, or any other entity. All fields below stay editable."
        ),
    )
    if selected == _CUSTOM_DEFENDANT:
        d = {
            "name": "", "address": "", "city": "", "state": "CA", "zip": "",
            "agent_name": "City Clerk", "agent_title": "City Clerk",
            "agent_address": "", "agent_city": "", "agent_state": "CA", "agent_zip": "",
        }
        kp = f"{key_prefix}_def{def_id}_custom"
    else:
        d = defendant_info(selected)
        # Selection is part of the key so switching cities refreshes the defaults
        kp = f"{key_prefix}_def{def_id}_{selected}"

    c1, c2 = st.columns(2)
    with c1:
        name_v   = st.text_input("Defendant Name *", value=d["name"], key=f"{kp}_name")
        street_v = st.text_input("Street Address", value=d["address"], key=f"{kp}_street")
        city_v   = st.text_input("City", value=d["city"], key=f"{kp}_city")
        s1, s2 = st.columns(2)
        with s1:
            state_v = st.text_input("State", value=d.get("state", "CA"), key=f"{kp}_state")
        with s2:
            zip_v = st.text_input("ZIP", value=d["zip"], key=f"{kp}_zip")

    out = {
        "name":    name_v.strip(),
        "address": street_v.strip(),
        "street":  street_v.strip(),
        "city":    city_v.strip(),
        "state":   state_v.strip() or "CA",
        "zip":     zip_v.strip(),
    }

    with c2:
        if is_primary:
            out["agent_name"]    = st.text_input("Agent for Service (Name)", value=d["agent_name"], key=f"{kp}_agent_name").strip() or "City Clerk"
            out["agent_title"]   = st.text_input("Agent Title", value=d["agent_title"], key=f"{kp}_agent_title").strip() or "City Clerk"
            out["agent_address"] = st.text_input("Agent Street", value=d["agent_address"], key=f"{kp}_agent_street").strip()
            out["agent_city"]    = st.text_input("Agent City", value=d["agent_city"], key=f"{kp}_agent_city").strip()
            out["agent_state"]   = "CA"
            out["agent_zip"]     = st.text_input("Agent ZIP", value=d["agent_zip"], key=f"{kp}_agent_zip").strip()
        else:
            out["phone"]     = st.text_input("Phone", key=f"{kp}_phone").strip()
            out["mailing"]   = st.text_input("Mailing Address (if different)", key=f"{kp}_mailing").strip()
            out["job_title"] = st.text_input("Job Title (if known)", key=f"{kp}_job_title").strip()

    if out["name"]:
        st.caption(f"**{out['name']}** · {out['address']}, {out['city']}, {out['state']} {out['zip']}")
    return out


tab_manual, tab_sheet = st.tabs(["📝 Manual Entry", "📊 Spreadsheet Import"])


# ══════════════════════════════════════════════════════
# TAB 1 — MANUAL ENTRY
# ══════════════════════════════════════════════════════

with tab_manual:
    # ════════════════════════════════════════════════════
    # STEP 1 — GOVERNMENT CLAIM (file this first)
    # ════════════════════════════════════════════════════
    st.header("Step 1 — Government Claim")
    st.caption(
        "Before you can sue a California city or public entity for property "
        "destroyed in a sweep, you must first file a government tort claim "
        "(Gov. Code §§ 905, 910) with that entity — generally within six months "
        "of the incident. Fill in the defendant, your information, and the "
        "incident below, then generate the claim form to file with the City Clerk."
    )

    # ── Defendants (dynamic list; extras go on SC-100A) ────────────────
    if "manual_def_ids" not in st.session_state:
        st.session_state["manual_def_ids"] = [0]
        st.session_state["manual_def_next"] = 1

    hdr_l, hdr_r = st.columns([0.92, 0.08])
    with hdr_l:
        st.subheader("Defendant")
    with hdr_r:
        if st.button("➕", key="manual_add_def", help="Add another defendant (listed on form SC-100A)"):
            st.session_state["manual_def_ids"].append(st.session_state["manual_def_next"])
            st.session_state["manual_def_next"] += 1
            st.rerun()

    manual_defendants = []
    for pos, def_id in enumerate(st.session_state["manual_def_ids"]):
        is_primary = pos == 0
        if is_primary:
            if len(st.session_state["manual_def_ids"]) > 1:
                st.markdown("**Defendant 1** · named on SC-100")
        else:
            rc1, rc2 = st.columns([0.92, 0.08])
            with rc1:
                st.markdown(f"**Defendant {pos + 1}** · on attached SC-100A")
            with rc2:
                if st.button("✕", key=f"manual_rm_def{def_id}", help="Remove this defendant"):
                    st.session_state["manual_def_ids"].remove(def_id)
                    st.rerun()
        manual_defendants.append(_defendant_block("manual", def_id, is_primary))
    st.divider()

    # ── Plaintiff ──────────────────────────────────────────────────────
    st.subheader("Plaintiff")
    c1, c2 = st.columns(2)
    with c1:
        name   = st.text_input("Full Legal Name *", placeholder="Jane Doe")
        street = st.text_input(
            "Street / Mailing Address",
            placeholder="c/o 1234 Telegraph Ave  (use c/o for unhoused clients)",
        )
        phone  = st.text_input("Phone", placeholder="510-555-0100")
    with c2:
        city = st.text_input("City", value="Oakland")
        cs1, cs2 = st.columns(2)
        with cs1:
            state = st.text_input("State", value="CA")
        with cs2:
            zip_  = st.text_input("ZIP", placeholder="94609")
        email = st.text_input("Email (optional)", placeholder="")

    # ── Incident & Claim (shared by the claim form and the lawsuit) ────
    st.divider()
    st.subheader("Incident & Claim")
    c1, c2 = st.columns(2)
    with c1:
        incident_date     = st.text_input("Date of Sweep *", placeholder="MM/DD/YYYY")
        incident_location = st.text_input(
            "Location of Sweep",
            placeholder="E.g. E 12th St & 16th Ave underpass, Oakland",
        )
    with c2:
        claim_amount = st.text_input("Claim Amount ($) *", placeholder="10000")
        involved_employees = st.text_input(
            "City employees or agencies involved (if known)",
            placeholder="DPW crew, police officers, contractor…",
        )

    claim_reason = st.text_area(
        "Brief summary of what happened (used on the claim form and SC-100)",
        placeholder=(
            "On [date], the City of Oakland DPW conducted an encampment sweep "
            "at [location] and destroyed Plaintiff's personal property…"
        ),
        height=120,
    )

    # ── Generate the government claim form ─────────────────────────────
    if st.button(
        "Generate Government Claim Form", type="primary",
        use_container_width=True, key="gen_govt_claim",
    ):
        _gc_def = manual_defendants[0] if manual_defendants else {}
        _clerk_addr = ", ".join(part for part in [
            (_gc_def.get("agent_address") or "").strip(),
            (_gc_def.get("agent_city") or "").strip(),
            f"CA {(_gc_def.get('agent_zip') or '').strip()}".strip(),
        ] if part)
        try:
            st.session_state["govt_claim_bytes"] = _build_govt_claim_docx({
                "entity":            _gc_def.get("name", ""),
                "clerk_address":     _clerk_addr,
                "claimant_name":     name.strip(),
                "claimant_address":  ", ".join(p for p in [
                    street.strip(), city.strip(),
                    f"{state.strip()} {zip_.strip()}".strip(),
                ] if p),
                "claimant_phone":    phone.strip(),
                "incident_date":     incident_date.strip(),
                "incident_location": incident_location.strip(),
                "description":       claim_reason.strip(),
                "employees":         involved_employees.strip(),
                "amount":            claim_amount.strip(),
            })
            st.session_state["govt_claim_name"] = (
                f"{_slug(name.strip() or 'claim')}_government_claim.docx"
            )
        except Exception as e:
            st.session_state.pop("govt_claim_bytes", None)
            st.error(f"Could not generate the claim form: {e}")
    if st.session_state.get("govt_claim_bytes"):
        st.download_button(
            "⬇️ Download Government Claim (Word)",
            data=st.session_state["govt_claim_bytes"],
            file_name=st.session_state.get("govt_claim_name", "government_claim.docx"),
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
            key="dl_govt_claim",
        )
        st.caption(
            "Print, sign, and file this with the City Clerk (the agent for service "
            "shown in the defendant section). The city generally has **45 days** to "
            "respond. Once the claim is rejected — or 45 days pass — move to Step 2."
        )

    # ════════════════════════════════════════════════════
    # STEP 2 — FILE THE SMALL CLAIMS LAWSUIT
    # ════════════════════════════════════════════════════
    st.divider()
    st.header("Step 2 — File the Small Claims Lawsuit")
    st.caption(
        "After the claim is rejected (or 45 days pass with no response), file in "
        "small claims court. Pick the court, set your dates, itemize the property, "
        "and generate the filing packet: SC-100 (+ SC-100A), FW-001, FW-003, "
        "SC-112A, and SC-150."
    )

    # Court selector lives outside the form so county → courthouse cascade works
    st.subheader("Filing Court")
    manual_court = _court_selector("manual")
    st.caption(
        f"Court: **Superior Court of California, County of {manual_court['county']}** · "
        f"{manual_court['address']}, {manual_court['city']}, CA {manual_court['zip']}"
    )

    fd1, fd2 = st.columns(2)
    with fd1:
        filing_date = st.text_input("Filing Date *", placeholder="MM/DD/YYYY")
    with fd2:
        govt_claim_date = st.text_input(
            "Govt Claim Filed with City Clerk *", placeholder="MM/DD/YYYY",
            help="The date you filed (or will file) the Step 1 government claim.",
        )

    declaration_text_input = st.text_area(
        "Write your declaration in your own words",
        placeholder="Start typing what happened. For example: I was present when the City took my belongings, I was not given notice, and I saw them throw away my property.",
        height=180,
    )

    # ── Items ──────────────────────────────────────────────────────────
    st.divider()
    st.subheader("Itemized Property")
    st.caption("List each item that was destroyed, its estimated value, and its condition before the loss.")
    items_df = st.data_editor(
        pd.DataFrame({
            "Description": ["", "", ""],
            "Value ($)": ["", "", ""],
            "Condition": ["New", "Good", "Fair"],
        }),
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Description": st.column_config.TextColumn(width="large"),
            "Value ($)": st.column_config.TextColumn(width="small"),
            "Condition": st.column_config.SelectboxColumn(
                width="small",
                options=["New", "Good", "Fair", "Salvage"],
            ),
        },
        hide_index=True,
    )

    if st.button("Generate declaration", use_container_width=True):
        items = []
        for _, r in items_df.iterrows():
            description = str(r.get("Description", "")).strip()
            if not description:
                continue
            items.append({
                "description": description,
                "value": str(r.get("Value ($)", "")).strip(),
                "condition": str(r.get("Condition", "")).strip() or "Unknown",
            })

        declaration_text = _build_guided_declaration(
            declaration_text_input,
            {
                "incident_date": incident_date.strip(),
                "claim_amount": claim_amount.strip(),
                "items": items,
            },
        )
        st.session_state["declaration_text"] = declaration_text

    declaration_text = st.session_state.get("declaration_text", "")
    st.text_area(
        "Declaration draft",
        value=declaration_text or "Press the button above to generate a court-style declaration.",
        height=260,
    )
    damages_calc = st.text_area(
        "How Damages Are Calculated",
        placeholder=(
            "Itemize property value + emotional distress. "
            "Leave blank to auto-fill from description above."
        ),
        height=80,
    )

    # ── Fee Waiver ─────────────────────────────────────────────────────
    st.divider()
    st.subheader("Fee Waiver")
    c1, c2, c3 = st.columns(3)
    with c1:
        fw_basis = st.radio(
            "Basis",
            ["5c — Cannot afford fees", "5a — Public benefits", "5b — Income below threshold"],
            help="5c is correct for most encampment sweep clients.",
        )
        st.markdown("**Public benefits:**")
        recv_medi_cal = st.checkbox("Medi-Cal")
        recv_snap     = st.checkbox("CalFresh / SNAP")
        recv_calworks = st.checkbox("CalWORKS")
    with c2:
        income_source = st.text_input("Income Source", placeholder="General Assistance, SSI…")
        income_amount = st.text_input("Monthly Income ($)", placeholder="400")
        total_income  = st.text_input("Total Monthly Income ($)", placeholder="400")
    with c3:
        exp_food      = st.text_input("Food / Supplies ($)", value="0")
        exp_medical   = st.text_input("Medical / Dental ($)", value="0")
        exp_transport = st.text_input("Transportation ($)", value="0")
        exp_housing   = st.text_input("Housing ($)", value="0")
        total_expenses = st.text_input("Total Monthly Expenses ($)", placeholder="300")

    submitted = st.button(
        "Generate Forms", type="primary", use_container_width=True
    )

    st.download_button(
        "⬇️ Download declaration as Word document",
        data=_build_declaration_docx(declaration_text),
        file_name="guided_declaration.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        disabled=not declaration_text.strip(),
    )

    # ── Handle submission ──────────────────────────────────────────────────
    if submitted:
        items = []
        for _, r in items_df.iterrows():
            description = str(r.get("Description", "")).strip()
            if not description:
                continue
            items.append({
                "description": description,
                "value": str(r.get("Value ($)", "")).strip(),
                "condition": str(r.get("Condition", "")).strip() or "Unknown",
            })
        basis_code = fw_basis.split(" — ")[0].strip()

        _defendant = manual_defendants[0]

        case = {
            "court": manual_court,
            "plaintiff": {
                "name":   name.strip(),
                "street": street.strip(),
                "city":   city.strip(),
                "state":  state.strip(),
                "zip":    zip_.strip(),
                "phone":  phone.strip(),
                "email":  email.strip(),
            },
            "defendant": _defendant,
            "claim": {
                "amount":                claim_amount.strip(),
                "reason":                claim_reason.strip(),
                "incident_date":         incident_date.strip(),
                "damages_calculation":   damages_calc.strip() or claim_reason.strip(),
                "govt_claim_filed_date": govt_claim_date.strip(),
                "items":                 items,
            },
            "filing": {
                "filing_date":      filing_date.strip(),
                "demanded_payment": True,
            },
            "fee_waiver": {
                "basis":                  basis_code,
                "waive_option":           "all",
                "receives_medi_cal":      recv_medi_cal,
                "receives_snap":          recv_snap,
                "receives_calworks":      recv_calworks,
                "income_source_1":        income_source.strip(),
                "income_amount_1":        income_amount.strip(),
                "total_monthly_income":   total_income.strip(),
                "expense_housing":        exp_housing.strip(),
                "expense_food":           exp_food.strip(),
                "expense_utilities":      "0",
                "expense_medical":        exp_medical.strip(),
                "expense_transport":      exp_transport.strip(),
                "total_monthly_expenses": total_expenses.strip(),
            },
            "declaration": {
                "declarant_name": name.strip(),
                "content":        declaration_text.strip() or claim_reason.strip(),
            },
            "subpoena": {
                "case_caption":     f"{name.strip()} v. {_defendant.get('name') or 'City of Oakland'}",
                "to":               st.session_state.get("sub_recipient_name", "").strip(),
                "custodian":        st.session_state.get("sub_recipient_custodian", "").strip(),
                "service_location": st.session_state.get("sub_recipient_service", "").strip(),
                "requests": (
                    [r for _i, r in enumerate(_DEFAULT_SUBPOENA_REQUESTS)
                     if st.session_state.get(f"sub_req_{_i}", True)]
                    + [line.strip()
                       for line in st.session_state.get("sub_extra_requests", "").splitlines()
                       if line.strip()]
                )[:10],
            },
        }

        # Attach additional defendants (each generates a filled SC-100A)
        case['additional_defendants'] = [
            dd for dd in manual_defendants[1:] if dd.get('name')
        ]

        try:
            pdfs = _generate_pdfs(case)
            _show_downloads(pdfs, _slug(name.strip()))
            st.download_button(
                "💾  Save Case Data (JSON)",
                data=json.dumps(case, indent=2).encode(),
                file_name=f"{_slug(name.strip())}_case.json",
                mime="application/json",
            )
        except ValueError as e:
            st.error(str(e))
        except Exception as e:
            st.error(f"Unexpected error: {e}")


    # ════════════════════════════════════════════════════
    # STEP 3 — PREPARE FOR TRIAL
    # ════════════════════════════════════════════════════
    st.divider()
    st.header("Step 3 — Prepare for Trial")
    st.caption(
        "Once your case is filed, gather your evidence. Use the subpoena "
        "below to make the city or other agencies produce records — footage, "
        "reports, policies — before your hearing."
    )

    # ── Subpoena (SC-107) — separate, standalone form ──────────────────
    st.divider()
    st.subheader("Subpoena Request (SC-107)")
    st.caption(
        "This is its own form: generate the subpoena by itself with the button "
        "below, without generating the rest of the packet. The plaintiff and "
        "defendant at the top of the SC-107 fill in automatically from the "
        "sections above; enter who you're subpoenaing separately below. Filled "
        "fields are also included in the packet the next time you press "
        "**Generate Forms** in Step 2."
    )
    _default_requests = _DEFAULT_SUBPOENA_REQUESTS

    _sub_def = manual_defendants[0] if manual_defendants else {}
    _sub_def_name = (_sub_def.get("name") or "").strip()

    with st.container(border=True):
        # Case caption: auto-filled from the sections above, but shown as
        # real fields so they can be reviewed and edited before generating.
        st.markdown("**Case caption**")
        st.caption(
            "Auto-filled from the plaintiff and defendant sections above — "
            "edit either field if the SC-107 should read differently."
        )
        # Names are part of the keys so changes upstream refresh the defaults
        _cap_kp = f"subcap_{name.strip()}_{_sub_def_name}"
        cap1, cap2 = st.columns(2)
        with cap1:
            sub_plaintiff_name = st.text_input(
                "Plaintiff (top of SC-107)",
                value=name.strip(),
                key=f"{_cap_kp}_plaintiff",
            )
        with cap2:
            sub_defendant_name = st.text_input(
                "Defendant (top of SC-107)",
                value=_sub_def_name,
                key=f"{_cap_kp}_defendant",
            )

        # Separate contact block for the person/agency being subpoenaed
        st.markdown("**Who are you subpoenaing?**")
        st.caption(
            "Enter the contact information of the person or agency you want "
            "records from. This is separate from the defendant — it can be a "
            "police department, city agency, business, or individual."
        )
        sub_to = st.text_input(
            "Name of person or agency to subpoena",
            key="sub_recipient_name",
            placeholder="Oakland Police Department Records Division",
        )
        sub_custodian = st.text_input(
            "Custodian of records (person or division responsible for the records)",
            key="sub_recipient_custodian",
            placeholder="Records Division",
        )
        sub_service = st.text_input(
            "Service address (where the subpoena will be delivered)",
            key="sub_recipient_service",
            placeholder="1515 Clay St, Oakland CA 94612",
        )

        st.markdown("**Documents requested**")
        st.caption(
            "What documents would you like subpoenaed for your case? "
            "Check any that apply, and add your own below. Leave blank to skip."
        )
        subpoena_checks = {}
        for _i_req, req in enumerate(_default_requests):
            subpoena_checks[req] = st.checkbox(req, value=True, key=f"sub_req_{_i_req}")
        sub_extra = st.text_area(
            "Any additional documents to request (one per line)",
            key="sub_extra_requests",
            placeholder="Any other records or documents…",
            height=80,
        )

        if st.button(
            "Generate Subpoena Only (SC-107)",
            use_container_width=True,
            key="gen_sc107_only",
        ):
            _sub_case = {
                "plaintiff": {"name": sub_plaintiff_name.strip() or "Plaintiff"},
                "defendant": {**_sub_def, "name": sub_defendant_name.strip() or _sub_def_name},
                "case_number": "",
                "filing": {"filing_date": filing_date.strip()},
                "subpoena": {
                    "case_caption": f"{sub_plaintiff_name.strip() or 'Plaintiff'} v. {sub_defendant_name.strip() or 'Defendant'}",
                    "to":               sub_to.strip(),
                    "custodian":        sub_custodian.strip(),
                    "service_location": sub_service.strip(),
                    "requests": (
                        [r for r, checked in subpoena_checks.items() if checked]
                        + [line.strip() for line in sub_extra.splitlines() if line.strip()]
                    )[:10],
                },
            }
            try:
                with tempfile.TemporaryDirectory() as _td, _quiet():
                    _sub_out = Path(_td) / "sc107.pdf"
                    fill_sc107(_sub_case, str(_TPL / "sc107.pdf"), str(_sub_out))
                    st.session_state["sc107_only_bytes"] = _sub_out.read_bytes()
                st.session_state["sc107_only_name"] = (
                    f"{_slug(name.strip() or 'subpoena')}_sc107.pdf"
                )
            except Exception as e:
                st.session_state.pop("sc107_only_bytes", None)
                st.error(f"Could not generate SC-107: {e}")

        if st.session_state.get("sc107_only_bytes"):
            st.download_button(
                "⬇️ Download SC-107 Subpoena",
                data=st.session_state["sc107_only_bytes"],
                file_name=st.session_state.get("sc107_only_name", "sc107.pdf"),
                mime="application/pdf",
                use_container_width=True,
                key="dl_sc107_only",
            )



# ══════════════════════════════════════════════════════
# TAB 2 — SPREADSHEET IMPORT
# ══════════════════════════════════════════════════════

with tab_sheet:
    st.subheader("Batch Import from Spreadsheet")

    # ── Template downloads ─────────────────────────────────────────────────
    c_info, c_tmpl = st.columns([3, 1])
    with c_info:
        st.info(
            "**Oakland intake format** (columns: Name, Address, Phone Number, "
            "Location of Injury, Date of Injury) is auto-detected from your Google Sheet. "
            "You'll set claim amount, filing dates, and fee waiver defaults that apply to "
            "every client in the batch."
        )
    with c_tmpl:
        st.download_button(
            "📥 Download Full Template CSV",
            data=_csv_template_bytes(),
            file_name="cases_template.csv",
            mime="text/csv",
            width="stretch",
            help="Use this template if you want to specify all fields per client.",
        )

    # ── Column reference ───────────────────────────────────────────────────
    with st.expander("View full template column reference"):
        col_df = pd.DataFrame(
            [(c[0], "✓" if c[2] else "", c[1], c[3]) for c in _TEMPLATE_COLS],
            columns=["Column", "Required", "Description", "Example"],
        )
        st.dataframe(col_df, use_container_width=True, hide_index=True)

    with st.expander("SC-107 subpoena checklist and spreadsheet fields"):
        subpoena_df = pd.DataFrame([
            ["subpoena_case_caption",   "Case caption for SC-107 subpoena"],
            ["subpoena_to",             "Subpoena recipient or agency to which records are directed"],
            ["subpoena_custodian",      "Custodian of records responsible for producing documents"],
            ["subpoena_service_location","Service address where the subpoena may be delivered"],
            ["subpoena_request_1",      "Body-worn camera and officer dashboard footage from the sweep"],
            ["subpoena_request_2",      "Incident reports, notes, and supplemental reports related to the sweep"],
            ["subpoena_request_3",      "Dispatch logs, radio transmissions, and 911/311 recordings"],
            ["subpoena_request_4",      "Officer complaint, investigation, and disciplinary records"],
            ["subpoena_request_5",      "Internal communications, emails, memos, and directives about encampment sweeps"],
            ["subpoena_request_6",      "Policies, training materials, use-of-force guidelines, and encampment protocols"],
            ["subpoena_request_7",      "Property seizure, storage, chain-of-custody, and disposal records"],
            ["subpoena_request_8",      "Surveillance camera and private video footage from the sweep location"],
            ["subpoena_request_9",      "Records of coordination between police, DPW, and other agencies"],
            ["subpoena_request_10",     "Logs, schedules, and written directives authorizing the sweeps"],
        ], columns=["Spreadsheet Column", "Request description"])
        st.dataframe(subpoena_df, use_container_width=True, hide_index=True)

    # ── Upload ─────────────────────────────────────────────────────────────
    uploaded = st.file_uploader(
        "Upload spreadsheet (CSV or XLSX)",
        type=["csv", "xlsx"],
        label_visibility="collapsed",
    )

    if not uploaded:
        st.stop()

    try:
        if uploaded.name.endswith(".xlsx"):
            df = pd.read_excel(uploaded, dtype=str)
        else:
            df = pd.read_csv(uploaded, dtype=str)
        df = df.fillna("")
        # Drop rows where Name/plaintiff_name is blank
        name_col = "Name" if "Name" in df.columns else "plaintiff_name"
        df = df[df[name_col].str.strip() != ""].reset_index(drop=True)
    except Exception as e:
        st.error(f"Could not read file: {e}")
        st.stop()

    fmt = _detect_format(df)
    st.write(f"**{len(df)} client(s) found** — format: `{fmt}`")

    # ── Preview ────────────────────────────────────────────────────────────
    preview_cols = {
        "oakland_intake": ["Name", "Address", "Phone Number", "Location of Injury", "Date of Injury"],
        "template": [c[0] for c in _TEMPLATE_COLS if c[2]],  # required cols only
    }
    show_cols = [c for c in preview_cols.get(fmt, list(df.columns)) if c in df.columns]
    st.dataframe(df[show_cols] if show_cols else df, use_container_width=True, height=250)

    if fmt == "unknown":
        st.warning(
            "Column names not recognized. Rename columns to match the Oakland intake format "
            "(Name, Address, Phone Number, Location of Injury, Date of Injury) "
            "or the full template format (plaintiff_name, etc.)."
        )
        st.stop()

    # ═══════════════════════════════════════════════════════
    # OAKLAND INTAKE FORMAT: batch defaults form
    # ═══════════════════════════════════════════════════════
    if fmt == "oakland_intake":
        st.divider()
        st.subheader("Batch Settings")
        st.caption(
            "These values apply to **all clients** in the spreadsheet. "
            "They fill in the fields not captured in the Oakland intake sheet."
        )

        st.markdown("**Filing Court**")
        batch_court = _court_selector("batch")
        st.caption(
            f"Court: **Superior Court of California, County of {batch_court['county']}** · "
            f"{batch_court['address']}, {batch_court['city']}, CA {batch_court['zip']}"
        )

        st.markdown("**Defendant**")
        batch_def = _defendant_selector("batch")

        with st.form("batch_defaults_form", border=True):
            d1, d2 = st.columns(2)
            with d1:
                b_filing_date      = st.text_input("Filing Date *", placeholder="MM/DD/YYYY",
                                                    help="Date you're filing the small claims paperwork.")
                b_govt_claim_date  = st.text_input("Govt Claim Filed with City Clerk *",
                                                    placeholder="MM/DD/YYYY",
                                                    help="Date the government tort claim was filed.")
                b_claim_amount     = st.text_input("Claim Amount ($) per client", value="10000",
                                                    help="Max $12,500 for individuals.")
            with d2:
                b_fw_basis = st.radio(
                    "Fee Waiver Basis",
                    ["5c — Cannot afford fees", "5a — Public benefits", "5b — Income threshold"],
                    horizontal=True,
                )
                b_recv_medi_cal = st.checkbox("All clients receive Medi-Cal")
                b_recv_snap     = st.checkbox("All clients receive CalFresh / SNAP")

            st.markdown("**Income & Expenses (applies to all — edit per-client if needed)**")
            e1, e2, e3, e4 = st.columns(4)
            with e1:
                b_income_source = st.text_input("Income Source", placeholder="General Assistance")
                b_income_amount = st.text_input("Monthly Income ($)", placeholder="400")
                b_total_income  = st.text_input("Total Monthly Income ($)", placeholder="400")
            with e2:
                b_exp_food    = st.text_input("Food ($)", value="0")
                b_exp_medical = st.text_input("Medical ($)", value="0")
            with e3:
                b_exp_transport = st.text_input("Transport ($)", value="0")
                b_exp_housing   = st.text_input("Housing ($)", value="0")
            with e4:
                b_total_expenses = st.text_input("Total Expenses ($)", placeholder="300")

            st.markdown("**Claim Narrative** (optional — leave blank for auto-generated text)")
            b_claim_reason = st.text_area(
                "Claim reason template",
                placeholder=(
                    "Leave blank to auto-generate: 'On [date], City of Oakland DPW "
                    "swept [location]…'  — incident date and location are filled "
                    "from each client's row."
                ),
                height=80,
            )
            b_declaration = st.text_area(
                "Declaration template",
                placeholder="Leave blank to auto-generate a first-person declaration.",
                height=80,
            )

            run_batch = st.form_submit_button(
                "Generate All Forms", type="primary", width="stretch"
            )

        if run_batch:
            defaults = {
                "court":                 batch_court,
                "defendant":             batch_def,
                "filing_date":           b_filing_date.strip(),
                "govt_claim_filed_date": b_govt_claim_date.strip(),
                "claim_amount":          b_claim_amount.strip(),
                "fw_basis":              b_fw_basis.split(" — ")[0].strip(),
                "receives_medi_cal":     b_recv_medi_cal,
                "receives_snap":         b_recv_snap,
                "receives_calworks":     False,
                "income_source":         b_income_source.strip(),
                "income_amount":         b_income_amount.strip(),
                "total_income":          b_total_income.strip(),
                "expense_food":          b_exp_food.strip(),
                "expense_medical":       b_exp_medical.strip(),
                "expense_transport":     b_exp_transport.strip(),
                "expense_housing":       b_exp_housing.strip(),
                "total_expenses":        b_total_expenses.strip(),
                "claim_reason":          b_claim_reason.strip(),
                "declaration":           b_declaration.strip(),
                "damages_calculation":   "",
            }

            results = []
            progress = st.progress(0, text="Generating forms…")
            for i, (_, row) in enumerate(df.iterrows()):
                pname = str(row.get("Name", f"Row {i+1}")).strip()
                try:
                    case = intake_row_to_case(row, defaults)
                    pdfs = _generate_pdfs(case)
                    results.append((pname, pdfs, None))
                except ValueError as e:
                    results.append((pname, None, str(e)))
                except Exception as e:
                    results.append((pname, None, f"Unexpected error: {e}"))
                progress.progress((i + 1) / len(df), text=f"Processed {i+1}/{len(df)}…")
            progress.empty()

            ok   = [(n, p, _) for n, p, _ in results if _ is None]
            fail = [(n, p, e) for n, p, e in results if e is not None]

            if ok:
                st.success(f"Generated forms for **{len(ok)}** of {len(results)} clients.")
                zip_buf = io.BytesIO()
                with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                    for pname, pdfs, _ in ok:
                        slug = _slug(pname)
                        for lbl, data in pdfs.items():
                            zf.writestr(
                                f"{slug}/{slug}_{lbl.lower().replace('-','')}.pdf", data
                            )
                st.download_button(
                    "⬇️  Download All Clients (ZIP)",
                    data=zip_buf.getvalue(),
                    file_name="oakland_encampment_forms.zip",
                    mime="application/zip",
                    type="primary",
                    width="stretch",
                )

            if fail:
                st.warning(f"{len(fail)} client(s) could not be processed:")
                for pname, _, err in fail:
                    with st.expander(f"Error — {pname}"):
                        st.text(err)

    # ═══════════════════════════════════════════════════════
    # TEMPLATE FORMAT: direct processing
    # ═══════════════════════════════════════════════════════
    elif fmt == "template":
        if st.button("Generate All Forms", type="primary", width="stretch"):
            results = []
            progress = st.progress(0, text="Generating forms…")
            for i, (_, row) in enumerate(df.iterrows()):
                pname = str(row.get("plaintiff_name", f"Row {i+1}")).strip()
                try:
                    case = template_row_to_case(row)
                    pdfs = _generate_pdfs(case)
                    results.append((pname, pdfs, None))
                except ValueError as e:
                    results.append((pname, None, str(e)))
                except Exception as e:
                    results.append((pname, None, f"Unexpected error: {e}"))
                progress.progress((i + 1) / len(df), text=f"Processed {i+1}/{len(df)}…")
            progress.empty()

            ok   = [(n, p, _) for n, p, _ in results if _ is None]
            fail = [(n, p, e) for n, p, e in results if e is not None]

            if ok:
                st.success(f"Generated forms for {len(ok)} of {len(results)} clients.")
                zip_buf = io.BytesIO()
                with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                    for pname, pdfs, _ in ok:
                        slug = _slug(pname)
                        for lbl, data in pdfs.items():
                            zf.writestr(
                                f"{slug}/{slug}_{lbl.lower().replace('-','')}.pdf", data
                            )
                st.download_button(
                    "⬇️  Download All Clients (ZIP)",
                    data=zip_buf.getvalue(),
                    file_name="all_cases.zip",
                    mime="application/zip",
                    type="primary",
                    width="stretch",
                )

            if fail:
                st.warning(f"{len(fail)} client(s) had errors:")
                for pname, _, err in fail:
                    with st.expander(f"Error — {pname}"):
                        st.text(err)


