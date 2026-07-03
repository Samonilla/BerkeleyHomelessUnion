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

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))

from fill_forms import (
    fill_sc100, fill_fw001, fill_fw003, fill_sc112a, fill_sc150,
    fill_sc105, fill_sc107, validate_case, DEFENDANT_DEFAULTS,
)
from smallclaims.fill_forms import fill_sc100a_for_party

_META_SC100 = str(HERE / "field_meta" / "sc100_fields.json")
_META_FW001 = str(HERE / "field_meta" / "fw001_fields.json")
_TPL = HERE / "templates"


# ─── PDF helpers ──────────────────────────────────────────────────────────────

def _slug(name: str) -> str:
    return re.sub(r"[^a-z0-9]+", "_", name.lower()).strip("_") or "case"


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
        "plaintiff": {
            "name":   g("Name"),
            "street": addr["street"],
            "city":   addr["city"],
            "state":  addr["state"],
            "zip":    addr["zip"],
            "phone":  phone,
            "email":  g("email"),
        },
        "defendant": DEFENDANT_DEFAULTS["city_of_oakland"],
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
        "defendant": DEFENDANT_DEFAULTS["city_of_oakland"],
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
    ("damages_calculation",     "Itemized damages breakdown",                     False, "Clothing $500..."),
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
    page_title="Oakland Small Claims Autofiller",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.title("Oakland Encampment — Small Claims Autofiller")
st.caption(
    "Generates SC-100, FW-001, FW-003, SC-112A, and SC-150 "
    "for City of Oakland encampment property destruction cases."
)

tab_manual, tab_sheet = st.tabs(["📝 Manual Entry", "📊 Spreadsheet Import"])


# ══════════════════════════════════════════════════════
# TAB 1 — MANUAL ENTRY
# ══════════════════════════════════════════════════════

with tab_manual:
    with st.form("manual_form", border=False):

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

        # ── Incident ───────────────────────────────────────────────────────
        st.divider()
        st.subheader("Incident & Claim")
        c1, c2 = st.columns(2)
        with c1:
            incident_date   = st.text_input("Date of Sweep *", placeholder="MM/DD/YYYY")
            filing_date     = st.text_input("Filing Date *", placeholder="MM/DD/YYYY")
        with c2:
            govt_claim_date = st.text_input(
                "Govt Claim Filed with City Clerk *", placeholder="MM/DD/YYYY"
            )
            claim_amount = st.text_input("Claim Amount ($) *", placeholder="10000")

        claim_reason = st.text_area(
            "What Happened *  (used on SC-100 and SC-150)",
            placeholder=(
                "On [date], the City of Oakland DPW conducted an encampment sweep "
                "at [location] and destroyed Plaintiff's personal property…"
            ),
            height=120,
        )
        damages_calc = st.text_area(
            "How Damages Are Calculated",
            placeholder=(
                "Itemize property value + emotional distress. "
                "Leave blank to auto-fill from description above."
            ),
            height=80,
        )

        # ── Items ──────────────────────────────────────────────────────────
        st.divider()
        st.subheader("Itemized Property (SC-112A attachment)")
        items_df = st.data_editor(
            pd.DataFrame({"Description": ["", "", ""], "Value ($)": ["", "", ""]}),
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "Description": st.column_config.TextColumn(width="large"),
                "Value ($)":   st.column_config.TextColumn(width="small"),
            },
            hide_index=True,
        )

        # ── Additional Defendants (SC-100A) ─────────────────────────────────
        st.divider()
        st.subheader("Additional Defendants (optional)")
        defs_df = st.data_editor(
            pd.DataFrame({"Name": [""], "Street": [""], "City": [""], "State": ["CA"], "ZIP": [""], "Phone": [""], "Mailing": [""], "Job Title": [""]}),
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
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

        # ── Declaration ────────────────────────────────────────────────────
        st.divider()
        st.subheader("Declaration (SC-150)")
        st.caption("Leave blank to use the incident description above.")
        declaration = st.text_area(
            "First-person statement (under penalty of perjury)",
            placeholder=(
                "I am the plaintiff in this action. On [date], the City of Oakland DPW…"
            ),
            height=120,
        )

        submitted = st.form_submit_button(
            "Generate Forms", type="primary", width="stretch"
        )

    # ── Handle submission ──────────────────────────────────────────────────
    if submitted:
        items = [
            {"description": str(r["Description"]).strip(), "value": str(r["Value ($)"]).strip()}
            for _, r in items_df.iterrows()
            if str(r["Description"]).strip()
        ]
        basis_code = fw_basis.split(" — ")[0].strip()

        case = {
            "plaintiff": {
                "name":   name.strip(),
                "street": street.strip(),
                "city":   city.strip(),
                "state":  state.strip(),
                "zip":    zip_.strip(),
                "phone":  phone.strip(),
                "email":  email.strip(),
            },
            "defendant": DEFENDANT_DEFAULTS["city_of_oakland"],
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
                "content":        declaration.strip() or claim_reason.strip(),
            },
            "subpoena": {
                "case_caption":     "",
                "to":               "",
                "custodian":        "",
                "service_location": "",
                "requests":         ["", "", "", "", "", "", "", "", "", ""],
            },
        }

        # Attach additional defendants from data editor
        additional_defendants = []
        for _, r in defs_df.iterrows():
            name = str(r.get('Name','')).strip()
            if not name:
                continue
            additional_defendants.append({
                'name': name,
                'street': str(r.get('Street','')).strip(),
                'city': str(r.get('City','')).strip(),
                'state': str(r.get('State','')).strip(),
                'zip': str(r.get('ZIP','')).strip(),
                'phone': str(r.get('Phone','')).strip(),
                'mailing': str(r.get('Mailing','')).strip(),
                'job_title': str(r.get('Job Title','')).strip(),
            })

        case['additional_defendants'] = additional_defendants

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

with tab_sc107:
    st.subheader("SC-107 Subpoena Helper")
    st.markdown(
        "Use this section to generate SC-107 subpoena request language and to download "
        "a helper CSV for subpoenas in lawsuits against police for destroying homeless "
        "encampments."
    )

    st.markdown("#### SC-107 spreadsheet helper fields")
    sc107_df = pd.DataFrame(
        _SC107_TEMPLATE_COLS,
        columns=["Spreadsheet Column", "Description", "Example"],
    )
    st.dataframe(sc107_df, use_container_width=True, hide_index=True)

    st.download_button(
        "📥 Download SC-107 helper CSV",
        data=_sc107_csv_template_bytes(),
        file_name="sc107_subpoena_helper.csv",
        mime="text/csv",
        type="secondary",
        width="stretch",
    )

    st.markdown("#### Suggested subpoena request categories")
    st.markdown(
        "- Body-worn camera, dashboard camera, and other video/audio recordings "
        "from the sweep.\n"
        "- Incident reports, field notes, supplemental reports, and emails by officers.\n"
        "- Dispatch logs, radio transmissions, and 911/311 call recordings.\n"
        "- Complaints, investigations, internal affairs files, and disciplinary history for involved officers.\n"
        "- Policies, training materials, written directives, and encampment sweep protocols.\n"
        "- Property seizure, storage, chain-of-custody, and disposal records.\n"
        "- All internal communications, memos, and coordination records between police, DPW, and other City agencies."
    )
