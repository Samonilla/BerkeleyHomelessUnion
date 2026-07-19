import io

import pandas as pd


INTAKE_COLS = {"Name", "Address", "Phone Number", "Location of Injury", "Date of Injury"}
TEMPLATE_COL = "plaintiff_name"


TEMPLATE_COLS = [
    ("plaintiff_name",          "Full legal name",                                True,  "Jane Doe"),
    ("plaintiff_street",        "Street / PO Box (c/o ... for unhoused)",         True,  "c/o 1234 Telegraph Ave"),
    ("incident_date",           "Date of sweep MM/DD/YYYY",                       True,  "05/12/2025"),
    ("claim_amount",            "Total claim dollars (max 12500)",                True,  "10000"),
    ("claim_reason",            "What happened - used on SC-100 and SC-150",      True,  "On May 12 2025 City of Oakland DPW..."),
    ("govt_claim_filed_date",   "Date govt tort claim filed with City Clerk",     True,  "08/15/2025"),
    ("filing_date",             "Date filing court papers MM/DD/YYYY",            True,  "09/15/2025"),
    ("total_monthly_income",    "Total monthly income $",                          True,  "400"),
    ("total_monthly_expenses",  "Total monthly expenses $",                        True,  "300"),
    ("plaintiff_city",          "Plaintiff's city",                               False, "Oakland"),
    ("plaintiff_state",         "State (default CA)",                             False, "CA"),
    ("plaintiff_zip",           "ZIP code",                                       False, "94609"),
    ("plaintiff_phone",         "Phone number",                                   False, "510-555-0100"),
    ("plaintiff_email",         "Email",                                          False, ""),
    ("damages_calculation",     "How damages were calculated (SC-100)",           False, "Clothing $500..."),
    ("income_source_1",         "Primary income source",                          False, "General Assistance"),
    ("income_amount_1",         "Primary income amount $",                         False, "400"),
    ("expense_food",            "Monthly food/supplies $",                         False, "200"),
    ("expense_medical",         "Monthly medical $",                               False, "50"),
    ("expense_transport",       "Monthly transport $",                             False, "50"),
    ("expense_housing",         "Monthly housing $",                               False, "0"),
    ("receives_medi_cal",       "Receives Medi-Cal? TRUE/FALSE",                  False, "TRUE"),
    ("fee_waiver_basis",        "Fee waiver basis: 5a 5b or 5c",                  False, "5c"),
    ("declaration_content",     "First-person declaration (optional)",            False, ""),
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
    ("subpoena_good_cause",     "Attachment 3: why good cause exists (blank = default)",   False, ""),
    ("subpoena_materiality",    "Attachment 4: why records are material (blank = default)", False, ""),
    ("item_1_desc",             "Property item 1 description",                    False, "Tent and sleeping bag"),
    ("item_1_value",            "Property item 1 value $",                         False, "350"),
    ("item_2_desc",             "Property item 2 description",                    False, "Clothing"),
    ("item_2_value",            "Property item 2 value $",                         False, "500"),
]


SC107_TEMPLATE_COLS = [
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
    ("subpoena_good_cause",      "Attachment 3: why good cause exists (blank = default)",   False, ""),
    ("subpoena_materiality",     "Attachment 4: why records are material (blank = default)", False, ""),
]


def detect_format(df: pd.DataFrame) -> str:
    cols = set(df.columns)
    if INTAKE_COLS.issubset(cols):
        return "oakland_intake"
    if TEMPLATE_COL in cols:
        return "template"
    return "unknown"


def csv_template_bytes() -> bytes:
    row = {c[0]: c[3] for c in TEMPLATE_COLS}
    buf = io.StringIO()
    pd.DataFrame([row]).to_csv(buf, index=False)
    return buf.getvalue().encode()


def sc107_csv_template_bytes() -> bytes:
    row = {c[0]: c[3] for c in SC107_TEMPLATE_COLS}
    buf = io.StringIO()
    pd.DataFrame([row]).to_csv(buf, index=False)
    return buf.getvalue().encode()
