"""
Berkeley Homeless Union — Officer Case Tracker (back office)

Internal interface for BHU officers to track small-claims cases captured by
the public intake app (app.py), update case status, keep notes, watch
deadlines, and export data.

Run:  streamlit run admin.py

Access: officer accounts (username + password). The first time the tracker
runs it asks you to create the admin account; signed-in officers can add
or remove accounts from the sidebar. Credentials are stored hashed
(PBKDF2) in admin_users.json — keep that file out of version control.
"""

import hashlib
import hmac
import io
import json
import os
import re
import secrets
import sys
import tempfile
import zipfile
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
import streamlit as st
from dateutil import parser as _dateutil

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))
from fill_forms import (  # noqa: E402
    fill_sc100, fill_fw001, fill_fw003, fill_sc112a, fill_sc150,
    fill_sc107, fill_sc100a_for_party,
)
from storage import (  # noqa: E402
    case_dirs as _case_dirs,
    primary_cases_dir as _primary_cases_dir,
    slug as _slug,
    normalize_org as _normalize_org,
    load_cases,
    save_case,
    import_case_json,
)

_TPL = HERE / "templates"
_META_SC100 = str(HERE / "field_meta" / "sc100_fields.json")
_META_FW001 = str(HERE / "field_meta" / "fw001_fields.json")

OUTCOMES = ["Pending", "Won", "Lost", "Settled", "Dismissed"]

STATUSES = [
    "Intake",
    "Govt Claim Filed",
    "Claim Rejected / 45 Days Passed",
    "Lawsuit Filed (SC-100)",
    "Trial Scheduled",
    "Resolved — Won",
    "Resolved — Settled",
    "Resolved — Lost",
    "Closed — Other",
]

GOVT_CLAIM_WINDOW_DAYS = 182   # ~6 months to present a govt tort claim
CITY_RESPONSE_DAYS = 45        # city's time to respond to the claim


# ─── Data loading / saving ────────────────────────────────────────────────────

def _parse_date(s):
    s = (str(s) if s is not None else "").strip()
    if not s:
        return None
    try:
        return _dateutil.parse(s, dayfirst=False).date()
    except Exception:
        return None


def generate_packet(case: dict):
    """Fill the lawsuit filing packet from a stored intake.

    Returns ({label: pdf_bytes}, [error strings]). Forms that fail are
    skipped and reported rather than aborting the whole packet.
    """
    pdfs, errors = {}, []
    with tempfile.TemporaryDirectory() as td:
        tmp = Path(td)
        jobs = [
            ("SC-100",  lambda o: fill_sc100(case, str(_TPL / "sc100.pdf"), o, _META_SC100)),
            ("FW-001",  lambda o: fill_fw001(case, str(_TPL / "fw001.pdf"), o, _META_FW001)),
            ("FW-003",  lambda o: fill_fw003(case, str(_TPL / "fw003.pdf"), o)),
            ("SC-112A", lambda o: fill_sc112a(case, str(_TPL / "sc112a.pdf"), o)),
            ("SC-150",  lambda o: fill_sc150(case, str(_TPL / "sc150.pdf"), o)),
        ]
        sub = case.get("subpoena") or {}
        if any(r for r in (sub.get("requests") or []) if r) or (sub.get("to") or "").strip():
            jobs.append(("SC-107", lambda o: fill_sc107(case, str(_TPL / "sc107.pdf"), o)))
        for label, fn in jobs:
            try:
                out = str(tmp / f"{label}.pdf")
                fn(out)
                pdfs[label] = Path(out).read_bytes()
            except Exception as e:
                errors.append(f"{label}: {e}")
        for i, ad in enumerate(case.get("additional_defendants") or [], start=1):
            try:
                out = str(tmp / f"sc100a_{i}.pdf")
                fill_sc100a_for_party(case, out, ad, role="defendant")
                pdfs[f"SC-100A-{i}"] = Path(out).read_bytes()
            except Exception as e:
                errors.append(f"SC-100A-{i}: {e}")
    return pdfs, errors


def packet_zip(pdfs: dict, slug: str) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for label, data in pdfs.items():
            zf.writestr(f"{slug}_{label.lower().replace('-', '')}.pdf", data)
    return buf.getvalue()


def deadline_info(case: dict) -> dict:
    """Compute deadline dates and flags for one case."""
    today = datetime.now().date()
    incident = _parse_date((case.get("claim") or {}).get("incident_date"))
    claim_filed = _parse_date((case.get("claim") or {}).get("govt_claim_filed_date"))
    status = (case.get("tracking") or {}).get("status", "Intake")

    info = {"claim_deadline": None, "response_due": None, "flag": "", "flag_detail": ""}

    if incident:
        info["claim_deadline"] = incident + timedelta(days=GOVT_CLAIM_WINDOW_DAYS)
    if claim_filed:
        info["response_due"] = claim_filed + timedelta(days=CITY_RESPONSE_DAYS)

    if status == "Intake" and info["claim_deadline"]:
        days_left = (info["claim_deadline"] - today).days
        if days_left < 0:
            info["flag"] = "🔴"
            info["flag_detail"] = f"Govt claim window passed {-days_left}d ago"
        elif days_left <= 30:
            info["flag"] = "🟠"
            info["flag_detail"] = f"{days_left}d left to file govt claim"
    elif status == "Govt Claim Filed" and info["response_due"]:
        days_over = (today - info["response_due"]).days
        if days_over >= 0:
            info["flag"] = "🟢"
            info["flag_detail"] = "45 days passed — can file lawsuit now"
        elif days_over >= -7:
            info["flag"] = "🟡"
            info["flag_detail"] = f"45-day mark in {-days_over}d"

    # Upcoming trial (any non-resolved case)
    if not info["flag"] and not str(status).startswith(("Resolved", "Closed")):
        trial = _parse_date((case.get("lawsuit") or {}).get("trial_date"))
        if trial:
            days = (trial - today).days
            if days == 0:
                info["flag"], info["flag_detail"] = "🔵", "Trial TODAY"
            elif 0 < days <= 14:
                info["flag"], info["flag_detail"] = "🔵", f"Trial in {days}d"
    return info


def flatten(path: Path, case: dict) -> dict:
    """One row per case for the table / export."""
    p = case.get("plaintiff") or {}
    c = case.get("claim") or {}
    d = case.get("defendant") or {}
    t = case.get("tracking") or {}
    lw = case.get("lawsuit") or {}
    dl = deadline_info(case)
    items = c.get("items") or []
    return {
        "Case #": case.get("internal_case_number", "(unassigned)"),
        "Name": p.get("name", ""),
        "Status": t.get("status", "Intake"),
        "⚑": dl["flag"],
        "Deadline note": dl["flag_detail"],
        "Court case #": case.get("case_number", ""),
        "Trial": lw.get("trial_date", ""),
        "Outcome": lw.get("outcome", ""),
        "Judgment $": lw.get("judgment_amount", ""),
        "Defendant": d.get("name", ""),
        "Incident": c.get("incident_date", ""),
        "Claim $": c.get("amount", ""),
        "Govt claim filed": c.get("govt_claim_filed_date", ""),
        "Claim deadline": str(dl["claim_deadline"] or ""),
        "45-day mark": str(dl["response_due"] or ""),
        "Items": len([i for i in items if (i.get("description") or "").strip()]),
        "Phone": p.get("phone", ""),
        "Captured": (case.get("captured_at") or "")[:10],
        "Last update": (t.get("updated_at") or "")[:10],
        "Officer notes": t.get("notes", ""),
        "_file": path.name,
    }


# ─── Officer accounts (shared with the intake site — see accounts.py) ─────────

from accounts import (  # noqa: E402
    USERS_FILE,
    add_user as _add_user,
    hash_password as _hash_password,
    load_users as _load_users,
    save_users as _save_users,
    user_org as _user_org,
    verify_login as _verify_login,
)


def _active_org() -> str:
    try:
        qp_org = st.query_params.get("org")
    except Exception:
        qp_org = ""
    env_org = os.environ.get("BHU_ORG", "")
    return _normalize_org(qp_org or env_org or "berkeley")


def _org_label(org: str) -> str:
    labels = {
        "berkeley": "Berkeley Homeless Union",
        "santa_rosa": "Santa Rosa Homeless Union",
    }
    return labels.get(org, org.replace("_", " ").title())


_ACTIVE_ORG = _active_org()
_ACTIVE_ORG_LABEL = _org_label(_ACTIVE_ORG)


# ─── Page & access gate ───────────────────────────────────────────────────────

st.set_page_config(page_title="BHU — Officer Case Tracker", page_icon="📋", layout="wide")

_users = _load_users()
_org_users = [u for u in sorted(_users) if _user_org(_users, u) == _ACTIVE_ORG]

if not _org_users:
    # First run: create the admin account before anything is shown.
    st.title(f"📋 {_ACTIVE_ORG_LABEL} — Officer Case Tracker Setup")
    st.caption(
        "No officer accounts exist yet for this union. Create the admin account to lock "
        "this tracker. You can add more officer accounts later from the "
        "sidebar."
    )
    with st.form("bhu_create_admin"):
        nu = st.text_input("Admin username")
        np1 = st.text_input("Password (min 8 characters)", type="password")
        np2 = st.text_input("Confirm password", type="password")
        if st.form_submit_button("Create admin account", type="primary"):
            if np1 != np2:
                st.error("Passwords don't match.")
            else:
                err = _add_user(_users, nu, np1, org=_ACTIVE_ORG)
                if err:
                    st.error(err)
                else:
                    st.session_state["bhu_user"] = nu.strip().lower()
                    st.rerun()
    st.stop()

if not st.session_state.get("bhu_user"):
    st.title(f"📋 {_ACTIVE_ORG_LABEL} — Officer Case Tracker")
    with st.form("bhu_login"):
        lu = st.text_input("Username")
        lp = st.text_input("Password", type="password")
        if st.form_submit_button("Sign in", type="primary"):
            if _verify_login(_users, lu, lp) and _user_org(_users, lu) == _ACTIVE_ORG:
                st.session_state["bhu_user"] = lu.strip().lower()
                st.rerun()
            else:
                st.error("Wrong username/password, or account belongs to another union.")
    st.stop()

# ─── Sidebar: session + account management ────────────────────────────────────

with st.sidebar:
    st.markdown(f"Union: **{_ACTIVE_ORG_LABEL}**")
    st.markdown(f"Signed in as **{st.session_state['bhu_user']}**")
    if st.button("Sign out", use_container_width=True):
        st.session_state.pop("bhu_user", None)
        st.rerun()

    with st.expander("Manage officer accounts"):
        org_accounts = [u for u in sorted(_users) if _user_org(_users, u) == _ACTIVE_ORG]
        st.caption(f"{len(org_accounts)} account(s) for this union: " + ", ".join(org_accounts))
        with st.form("bhu_add_officer", clear_on_submit=True):
            au = st.text_input("New officer username")
            ap = st.text_input("Password (min 8 characters)", type="password")
            if st.form_submit_button("Add account"):
                err = _add_user(_users, au, ap, org=_ACTIVE_ORG)
                st.error(err) if err else st.success(f"Added {au.strip().lower()}.")
        rm = st.selectbox(
            "Remove an account",
            [""] + [u for u in org_accounts if u != st.session_state["bhu_user"]],
        )
        if rm and st.button(f"Remove {rm}", use_container_width=True):
            _users.pop(rm, None)
            _save_users(_users)
            st.rerun()

    with st.expander("Change my password"):
        with st.form("bhu_change_pw", clear_on_submit=True):
            cur = st.text_input("Current password", type="password")
            new1 = st.text_input("New password (min 8 characters)", type="password")
            new2 = st.text_input("Confirm new password", type="password")
            if st.form_submit_button("Change password"):
                me = st.session_state["bhu_user"]
                if not _verify_login(_users, me, cur):
                    st.error("Current password is wrong.")
                elif new1 != new2:
                    st.error("New passwords don't match.")
                elif len(new1) < 8:
                    st.error("Password must be at least 8 characters.")
                else:
                    salt, digest = _hash_password(new1)
                    _users[me].update({"salt": salt, "hash": digest})
                    _save_users(_users)
                    st.success("Password changed.")

    with st.expander("Import case JSON files"):
        st.caption("Upload JSON files exported from the intake app (Save Case Data).")
        st.caption("Reading from: " + " | ".join(str(p) for p in _case_dirs()))
        files = st.file_uploader(
            "Case JSON files",
            type=["json"],
            accept_multiple_files=True,
            key="admin_case_import",
        )
        if files and st.button("Import uploaded files", use_container_width=True):
            ok_count, err_count = 0, 0
            for f in files:
                ok, msg = import_case_json(f.getvalue(), f.name, org=_ACTIVE_ORG)
                if ok:
                    ok_count += 1
                    st.success(msg)
                else:
                    err_count += 1
                    st.error(msg)
            st.info(f"Imported: {ok_count} · Errors: {err_count}")
            if ok_count:
                st.rerun()

st.title(f"📋 {_ACTIVE_ORG_LABEL} — Officer Case Tracker")
st.caption(
    "Every intake captured by the public site appears here under its internal "
    "case number (YYYYMMDD-initials). Update statuses, keep notes, watch "
    "deadlines, and export data. This page and the case files contain "
    "sensitive personal information — handle accordingly."
)

records = load_cases(org=_ACTIVE_ORG)
if not records:
    st.info("No cases captured yet for this union. Records appear here automatically once "
            "someone generates forms on the intake site.")
    st.stop()

rows = [flatten(p, c) for p, c in records]
df = pd.DataFrame(rows)

# ─── Overview metrics ─────────────────────────────────────────────────────────

m1, m2, m3, m4, m5 = st.columns(5)
m1.metric("Cases", len(df))
m2.metric("Active", int((~df["Status"].str.startswith(("Resolved", "Closed"))).sum()))
m3.metric("Flagged deadlines", int((df["⚑"] != "").sum()))
_amt = pd.to_numeric(
    df["Claim $"].astype(str).str.replace(r"[$,]", "", regex=True), errors="coerce"
)
m4.metric("Total claimed", f"${_amt.fillna(0).sum():,.0f}")
m5.metric("Resolved", int(df["Status"].str.startswith("Resolved").sum()))

# ─── Filters ──────────────────────────────────────────────────────────────────

st.divider()
f1, f2, f3 = st.columns([2, 2, 1])
with f1:
    q = st.text_input("Search (name, case #, defendant, notes)", placeholder="jane / 20260712-JD / berkeley…")
with f2:
    status_sel = st.multiselect("Status", STATUSES, default=[])
with f3:
    flagged_only = st.checkbox("⚑ Flagged only")

view = df
if q.strip():
    needle = q.strip().lower()
    mask = (
        view["Name"].str.lower().str.contains(needle, regex=False)
        | view["Case #"].str.lower().str.contains(needle, regex=False)
        | view["Defendant"].str.lower().str.contains(needle, regex=False)
        | view["Officer notes"].str.lower().str.contains(needle, regex=False)
    )
    view = view[mask]
if status_sel:
    view = view[view["Status"].isin(status_sel)]
if flagged_only:
    view = view[view["⚑"] != ""]

st.dataframe(
    view.drop(columns=["_file"]),
    use_container_width=True,
    hide_index=True,
)

# ─── Export ───────────────────────────────────────────────────────────────────

e1, e2, e3 = st.columns(3)
export_df = view.drop(columns=["_file"])
with e1:
    st.download_button(
        "⬇️ Export filtered view (CSV)",
        data=export_df.to_csv(index=False).encode(),
        file_name=f"bhu_cases_{datetime.now():%Y%m%d}.csv",
        mime="text/csv",
        use_container_width=True,
    )
with e2:
    _xbuf = io.BytesIO()
    export_df.to_excel(_xbuf, index=False, sheet_name="Cases")
    st.download_button(
        "⬇️ Export filtered view (Excel)",
        data=_xbuf.getvalue(),
        file_name=f"bhu_cases_{datetime.now():%Y%m%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
with e3:
    _zbuf = io.BytesIO()
    with zipfile.ZipFile(_zbuf, "w", zipfile.ZIP_DEFLATED) as zf:
        for path, _case in records:
            zf.write(path, arcname=path.name)
    st.download_button(
        "⬇️ Full records (JSON zip)",
        data=_zbuf.getvalue(),
        file_name=f"bhu_case_records_{datetime.now():%Y%m%d}.zip",
        mime="application/zip",
        use_container_width=True,
    )

# ─── Data capture audit ───────────────────────────────────────────────────────

st.divider()
with st.expander("🔍 Data capture audit — every field we collect"):
    st.caption(
        "A field-level look at what the intake site is capturing. Left: each "
        "field, how often it's filled, and an example value. Right: the full "
        "raw data, one row per case, one column per field."
    )

    def _cell(v):
        if isinstance(v, (list, dict)):
            return json.dumps(v, ensure_ascii=False)
        return v

    flat_full = pd.json_normalize([c for _, c in records], sep=".")
    flat_full = flat_full.apply(lambda col: col.map(_cell))

    def _filled(col):
        return col.map(
            lambda v: v is not None and str(v).strip() not in ("", "[]", "{}", "nan", "None")
        )

    _fill_mask = flat_full.apply(_filled)
    _examples = [
        (flat_full[c][_fill_mask[c]].astype(str).iloc[0][:60] if _fill_mask[c].any() else "")
        for c in flat_full.columns
    ]
    coverage = pd.DataFrame({
        "Field": flat_full.columns,
        "Filled": _fill_mask.sum().values,
        "% filled": (_fill_mask.mean() * 100).round(0).astype(int).astype(str) + "%",
        "Example": _examples,
    })

    st.markdown(
        f"Capturing **{flat_full.shape[1]} fields** across **{len(flat_full)} case(s)**."
    )
    a1, a2 = st.columns([2, 3])
    with a1:
        st.markdown("**Field coverage**")
        st.dataframe(coverage, use_container_width=True, hide_index=True, height=420)
    with a2:
        st.markdown("**Full captured data (all fields, all cases)**")
        st.dataframe(flat_full, use_container_width=True, height=420)

    st.download_button(
        "⬇️ Download ALL captured data (CSV — every field)",
        data=flat_full.to_csv(index=False).encode(),
        file_name=f"bhu_full_capture_{datetime.now():%Y%m%d}.csv",
        mime="text/csv",
        use_container_width=True,
    )

# ─── Case detail & tracking ───────────────────────────────────────────────────

st.divider()
st.header("Case Detail")

_labels = {
    f"{r['Case #']} — {r['Name']}  ({r['_file']})": r["_file"] for r in rows
}
sel_label = st.selectbox("Open a case", sorted(_labels))
sel_file = _labels[sel_label]
sel_path, sel_case = next((p, c) for p, c in records if p.name == sel_file)

p = sel_case.get("plaintiff") or {}
c = sel_case.get("claim") or {}
d = sel_case.get("defendant") or {}
t = sel_case.setdefault("tracking", {})
dl = deadline_info(sel_case)

left, right = st.columns([3, 2])

with left:
    st.subheader(f"{sel_case.get('internal_case_number', '(unassigned)')} — {p.get('name', '')}")
    if dl["flag"]:
        st.warning(f"{dl['flag']} {dl['flag_detail']}")
    st.markdown(
        f"**Defendant:** {d.get('name', '—')}  \n"
        f"**Incident:** {c.get('incident_date', '—')} · "
        f"**Claim:** ${c.get('amount', '—')} · "
        f"**Govt claim filed:** {c.get('govt_claim_filed_date') or '—'}  \n"
        f"**Govt-claim deadline:** {dl['claim_deadline'] or '—'} · "
        f"**45-day mark:** {dl['response_due'] or '—'}  \n"
        f"**Contact:** {p.get('phone') or '—'} · {p.get('email') or '—'} · "
        f"{p.get('street') or '—'}, {p.get('city') or ''}"
    )

    with st.expander("Itemized property (editable)"):
        def _s(v):
            return "" if pd.isna(v) else str(v).strip()

        _items_df = pd.DataFrame(
            [
                {
                    "Description": _s(i.get("description")),
                    "Value ($)": _s(i.get("value")),
                    "Condition": _s(i.get("condition")),
                }
                for i in (c.get("items") or [])
            ]
            or [{"Description": "", "Value ($)": "", "Condition": ""}]
        )
        _edited_items = st.data_editor(
            _items_df,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            key=f"items_{sel_file}",
        )
        if st.button("💾 Save items", key=f"items_save_{sel_file}"):
            new_items = []
            for _, r in _edited_items.iterrows():
                desc = _s(r.get("Description"))
                if not desc:
                    continue
                new_items.append({
                    "description": desc,
                    "value": _s(r.get("Value ($)")),
                    "condition": _s(r.get("Condition")),
                })
            sel_case.setdefault("claim", {})["items"] = new_items
            sel_case.setdefault("tracking", {}).setdefault("history", []).append(
                {"at": datetime.now().isoformat(timespec="seconds"),
                 "officer": st.session_state["bhu_user"],
                 "change": f"Itemized property updated ({len(new_items)} item(s))"}
            )
            try:
                save_case(sel_path, sel_case)
                st.success("Items saved.")
                st.rerun()
            except Exception as e:
                st.error(f"Could not save: {e}")
    with st.expander("Claim narrative"):
        st.write(c.get("reason") or "—")
    with st.expander("Declaration"):
        st.write((sel_case.get("declaration") or {}).get("content") or "—")
    with st.expander("Subpoena info"):
        st.json(sel_case.get("subpoena") or {})
    with st.expander("Everything (raw record)"):
        st.json(sel_case)

with right:
    st.subheader("Tracking")
    cur_status = t.get("status", "Intake")
    new_status = st.selectbox(
        "Status",
        STATUSES,
        index=STATUSES.index(cur_status) if cur_status in STATUSES else 0,
        key=f"status_{sel_file}",
    )
    new_notes = st.text_area(
        "Officer notes",
        value=t.get("notes", ""),
        height=180,
        key=f"notes_{sel_file}",
        placeholder="Contact attempts, hearing dates, evidence collected…",
    )
    officer = st.text_input("Your name / initials", value=t.get("last_officer", ""),
                            key=f"officer_{sel_file}")

    if st.button("💾 Save tracking update", use_container_width=True, type="primary"):
        now = datetime.now().isoformat(timespec="seconds")
        history = t.setdefault("history", [])
        if new_status != cur_status:
            history.append({"at": now, "officer": officer.strip(),
                            "change": f"Status: {cur_status} → {new_status}"})
        if new_notes != t.get("notes", ""):
            history.append({"at": now, "officer": officer.strip(), "change": "Notes updated"})
        t.update({
            "status": new_status,
            "notes": new_notes,
            "last_officer": officer.strip(),
            "updated_at": now,
        })
        if not sel_case.get("internal_case_number"):
            initials = "".join(
                w[0].upper() for w in re.split(r"\s+", (p.get("name") or "").strip())
                if w and w[0].isalpha()
            )
            sel_case["internal_case_number"] = f"{datetime.now():%Y%m%d}-{initials or 'XX'}"
        try:
            save_case(sel_path, sel_case)
            st.success("Saved.")
            st.rerun()
        except Exception as e:
            st.error(f"Could not save: {e}")

    if t.get("history"):
        st.markdown("**History**")
        for h in reversed(t["history"][-15:]):
            st.caption(f"{h.get('at', '')[:16]} · {h.get('officer') or '—'} · {h.get('change', '')}")

# ─── Edit intake data (fix mistakes from the website) ─────────────────────────

def _get_path(d: dict, path: str, default=""):
    cur = d
    for part in path.split("."):
        if not isinstance(cur, dict):
            return default
        cur = cur.get(part)
        if cur is None:
            return default
    return cur


def _set_path(d: dict, path: str, value) -> None:
    parts = path.split(".")
    cur = d
    for part in parts[:-1]:
        cur = cur.setdefault(part, {})
    cur[parts[-1]] = value


st.divider()
with st.expander("✏️ Edit intake data — fix mistakes in what was entered on the website"):
    st.caption(
        "Everything captured from the intake site is editable here. Saved "
        "changes go back into the case record and are used the next time "
        "the filing packet is generated. (Itemized property is edited in "
        "the expander above; status, notes, and lawsuit data below.)"
    )

    _K = f"edit_{sel_file}"

    def _ti(label, path, col=None, area=False, height=100):
        target = col if col is not None else st
        widget = target.text_area if area else target.text_input
        kwargs = {"height": height} if area else {}
        return path, widget(label, value=str(_get_path(sel_case, path)), key=f"{_K}_{path}", **kwargs)

    with st.form(f"{_K}_form"):
        st.markdown("**Plaintiff**")
        e1, e2, e3 = st.columns(3)
        _edits = [
            _ti("Name", "plaintiff.name", e1),
            _ti("Street / mailing address", "plaintiff.street", e2),
            _ti("Phone", "plaintiff.phone", e3),
            _ti("City", "plaintiff.city", e1),
            _ti("State", "plaintiff.state", e2),
            _ti("ZIP", "plaintiff.zip", e3),
            _ti("Email", "plaintiff.email", e1),
        ]

        st.markdown("**Defendant**")
        d1, d2, d3 = st.columns(3)
        _edits += [
            _ti("Defendant name", "defendant.name", d1),
            _ti("Street address", "defendant.address", d2),
            _ti("City", "defendant.city", d3),
            _ti("State", "defendant.state", d1),
            _ti("ZIP", "defendant.zip", d2),
            _ti("Agent for service (name)", "defendant.agent_name", d3),
            _ti("Agent title", "defendant.agent_title", d1),
            _ti("Agent street", "defendant.agent_address", d2),
            _ti("Agent city", "defendant.agent_city", d3),
            _ti("Agent ZIP", "defendant.agent_zip", d1),
        ]

        st.markdown("**Claim & dates**")
        c1_, c2_, c3_ = st.columns(3)
        _edits += [
            _ti("Claim amount ($)", "claim.amount", c1_),
            _ti("Incident date (MM/DD/YYYY)", "claim.incident_date", c2_),
            _ti("Govt claim filed (MM/DD/YYYY)", "claim.govt_claim_filed_date", c3_),
            _ti("Filing date (MM/DD/YYYY)", "filing.filing_date", c1_),
            _ti("What happened (claim reason)", "claim.reason", area=True, height=120),
            _ti("How damages were calculated", "claim.damages_calculation", area=True, height=90),
        ]

        st.markdown("**Declaration**")
        _edits += [
            _ti("Declarant name", "declaration.declarant_name"),
            _ti("Declaration text", "declaration.content", area=True, height=120),
        ]

        st.markdown("**Subpoena (SC-107)**")
        s1_, s2_, s3_ = st.columns(3)
        _edits += [
            _ti("Subpoena to (agency/person)", "subpoena.to", s1_),
            _ti("Custodian of records", "subpoena.custodian", s2_),
            _ti("Service address", "subpoena.service_location", s3_),
            _ti("Good cause (Attachment 3)", "subpoena.good_cause", area=True, height=90),
            _ti("Materiality (Attachment 4)", "subpoena.materiality", area=True, height=90),
        ]
        _req_text = st.text_area(
            "Records requested (one per line — Attachment 2a)",
            value="\n".join(r for r in (_get_path(sel_case, "subpoena.requests", []) or []) if r),
            height=140,
            key=f"{_K}_requests",
        )

        st.markdown("**Fee waiver**")
        f1_, f2_, f3_ = st.columns(3)
        _edits += [
            _ti("Basis (5a / 5b / 5c)", "fee_waiver.basis", f1_),
            _ti("Income source", "fee_waiver.income_source_1", f2_),
            _ti("Income amount ($)", "fee_waiver.income_amount_1", f3_),
            _ti("Total monthly income ($)", "fee_waiver.total_monthly_income", f1_),
            _ti("Total monthly expenses ($)", "fee_waiver.total_monthly_expenses", f2_),
            _ti("Housing ($)", "fee_waiver.expense_housing", f3_),
            _ti("Food ($)", "fee_waiver.expense_food", f1_),
            _ti("Medical ($)", "fee_waiver.expense_medical", f2_),
            _ti("Transport ($)", "fee_waiver.expense_transport", f3_),
        ]
        b1_, b2_, b3_ = st.columns(3)
        _recv_mc = b1_.checkbox("Receives Medi-Cal", value=bool(_get_path(sel_case, "fee_waiver.receives_medi_cal", False)), key=f"{_K}_mc")
        _recv_sn = b2_.checkbox("Receives CalFresh / SNAP", value=bool(_get_path(sel_case, "fee_waiver.receives_snap", False)), key=f"{_K}_sn")
        _recv_cw = b3_.checkbox("Receives CalWORKs", value=bool(_get_path(sel_case, "fee_waiver.receives_calworks", False)), key=f"{_K}_cw")

        if st.form_submit_button("💾 Save all corrections", type="primary", use_container_width=True):
            for _path, _val in _edits:
                _set_path(sel_case, _path, _val.strip() if isinstance(_val, str) else _val)
            _set_path(sel_case, "subpoena.requests",
                      [ln.strip() for ln in _req_text.splitlines() if ln.strip()])
            _set_path(sel_case, "fee_waiver.receives_medi_cal", _recv_mc)
            _set_path(sel_case, "fee_waiver.receives_snap", _recv_sn)
            _set_path(sel_case, "fee_waiver.receives_calworks", _recv_cw)
            sel_case.setdefault("tracking", {}).setdefault("history", []).append(
                {"at": datetime.now().isoformat(timespec="seconds"),
                 "officer": st.session_state["bhu_user"],
                 "change": "Intake data edited"}
            )
            try:
                save_case(sel_path, sel_case)
                st.success("Saved — regenerate the filing packet below to get corrected forms.")
                st.rerun()
            except Exception as e:
                st.error(f"Could not save: {e}")

# ─── Lawsuit: generate the filing packet & track court data ───────────────────

st.divider()
st.header("Lawsuit — Generate & Track")
lw = sel_case.setdefault("lawsuit", {})

lw_left, lw_right = st.columns([3, 2])

with lw_left:
    st.subheader("Court data")
    lc1, lc2 = st.columns(2)
    with lc1:
        lw_case_no = st.text_input(
            "Court case number (assigned at filing)",
            value=sel_case.get("case_number", ""),
            key=f"lw_caseno_{sel_file}",
        )
        lw_filed = st.text_input(
            "Lawsuit filed on (MM/DD/YYYY)",
            value=lw.get("filed_on", ""),
            key=f"lw_filed_{sel_file}",
        )
        lw_served = st.text_input(
            "Defendant served on (MM/DD/YYYY)",
            value=lw.get("served_on", ""),
            key=f"lw_served_{sel_file}",
        )
        lw_service = st.text_input(
            "Service method / server",
            value=lw.get("service_method", ""),
            placeholder="Sheriff / process server / certified mail…",
            key=f"lw_service_{sel_file}",
        )
    with lc2:
        lw_trial = st.text_input(
            "Trial date (MM/DD/YYYY)",
            value=lw.get("trial_date", ""),
            key=f"lw_trial_{sel_file}",
            help="Flags the case 🔵 in the table for the 14 days before trial.",
        )
        lw_dept = st.text_input(
            "Department / courtroom",
            value=lw.get("department", ""),
            key=f"lw_dept_{sel_file}",
        )
        _cur_outcome = lw.get("outcome", "Pending")
        lw_outcome = st.selectbox(
            "Outcome",
            OUTCOMES,
            index=OUTCOMES.index(_cur_outcome) if _cur_outcome in OUTCOMES else 0,
            key=f"lw_outcome_{sel_file}",
        )
        lw_judgment = st.text_input(
            "Judgment / settlement amount ($)",
            value=lw.get("judgment_amount", ""),
            key=f"lw_judgment_{sel_file}",
        )
    lw_notes = st.text_area(
        "Lawsuit notes (evidence status, continuances, collection…)",
        value=lw.get("notes", ""),
        height=90,
        key=f"lw_notes_{sel_file}",
    )

    if st.button("💾 Save lawsuit data", use_container_width=True, key=f"lw_save_{sel_file}"):
        now = datetime.now().isoformat(timespec="seconds")
        lw.update({
            "filed_on": lw_filed.strip(),
            "served_on": lw_served.strip(),
            "service_method": lw_service.strip(),
            "trial_date": lw_trial.strip(),
            "department": lw_dept.strip(),
            "outcome": lw_outcome,
            "judgment_amount": lw_judgment.strip(),
            "notes": lw_notes,
            "updated_at": now,
        })
        # The court case number goes on every generated form
        sel_case["case_number"] = lw_case_no.strip()
        sel_case.setdefault("tracking", {}).setdefault("history", []).append(
            {"at": now, "officer": st.session_state["bhu_user"], "change": "Lawsuit data updated"}
        )
        try:
            save_case(sel_path, sel_case)
            st.success("Saved.")
            st.rerun()
        except Exception as e:
            st.error(f"Could not save: {e}")

with lw_right:
    st.subheader("Filing packet")
    st.caption(
        "Regenerates the lawsuit forms (SC-100 + SC-100A, FW-001, FW-003, "
        "SC-112A, SC-150, and the SC-107 package if subpoena info was "
        "captured) from this case's stored intake — including the court "
        "case number saved on the left."
    )
    if st.button("📄 Generate filing packet", use_container_width=True,
                 type="primary", key=f"lw_gen_{sel_file}"):
        with st.spinner("Filling forms…"):
            pdfs, errors = generate_packet(sel_case)
        st.session_state[f"lw_pdfs_{sel_file}"] = pdfs
        st.session_state[f"lw_errors_{sel_file}"] = errors
        now = datetime.now().isoformat(timespec="seconds")
        sel_case.setdefault("tracking", {}).setdefault("history", []).append(
            {"at": now, "officer": st.session_state["bhu_user"],
             "change": f"Filing packet generated ({len(pdfs)} forms)"}
        )
        try:
            save_case(sel_path, sel_case)
        except Exception:
            pass

    _pdfs = st.session_state.get(f"lw_pdfs_{sel_file}") or {}
    _errs = st.session_state.get(f"lw_errors_{sel_file}") or []
    if _pdfs:
        _slug = re.sub(r"[^a-z0-9]+", "_", (p.get("name") or "case").lower()).strip("_")
        st.success(f"Generated {len(_pdfs)} form(s).")
        st.download_button(
            "⬇️ Download packet (ZIP)",
            data=packet_zip(_pdfs, _slug),
            file_name=f"{sel_case.get('internal_case_number', 'case')}_{_slug}_packet.zip",
            mime="application/zip",
            use_container_width=True,
            key=f"lw_zip_{sel_file}",
        )
        for _lbl, _data in _pdfs.items():
            st.download_button(
                f"⬇️ {_lbl}",
                data=_data,
                file_name=f"{_slug}_{_lbl.lower().replace('-', '')}.pdf",
                mime="application/pdf",
                use_container_width=True,
                key=f"lw_dl_{sel_file}_{_lbl}",
            )
    for _e in _errs:
        st.warning(_e)
