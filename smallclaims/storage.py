"""Shared case storage helpers for the small-claims apps.

Keep this file simple: one place for case-path rules, record loading,
record writing, and filename slugging.
"""

from __future__ import annotations

import json
import os
import re
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).parent
DEFAULT_CASES_DIR = HERE / "cases"
PERSISTENT_CASES_DIR = Path("/mount/data/bhu_cases")


def slug(name: str) -> str:
    return re.sub(r"[^a-z0-9]+", "_", (name or "").lower()).strip("_") or "case"


def normalize_org(org: str | None) -> str:
    token = re.sub(r"[^a-z0-9]+", "_", str(org or "").strip().lower()).strip("_")
    return token or "berkeley"


def case_org(case: dict) -> str:
    return normalize_org(
        case.get("organization")
        or case.get("org")
        or case.get("tenant")
        or "berkeley"
    )


def case_dirs() -> list[Path]:
    """Return case directories from most to least preferred."""
    env_dir = (os.environ.get("BHU_CASES_DIR") or "").strip()
    candidates = [
        Path(env_dir).expanduser() if env_dir else DEFAULT_CASES_DIR,
        PERSISTENT_CASES_DIR,
        DEFAULT_CASES_DIR,
        HERE.parent / "cases",
        Path("/mount/src/berkeleyhomelessunion/cases"),
        Path("/mount/src/berkeleyhomelessunion/smallclaims/cases"),
    ]
    ordered = []
    seen = set()
    for path in candidates:
        key = str(path)
        if key in seen:
            continue
        seen.add(key)
        ordered.append(path)
    return ordered


def primary_cases_dir() -> Path:
    return case_dirs()[0]


def load_cases(org: str | None = None) -> list[tuple[Path, dict]]:
    """Load all real case JSON files from the known storage locations."""
    by_key: dict[tuple[str, str], tuple[Path, dict]] = {}
    target_org = normalize_org(org) if org is not None else None
    for directory in case_dirs():
        if not directory.exists():
            continue
        for path in sorted(directory.glob("*.json")):
            if path.name.startswith(("sample", "_")):
                continue
            try:
                case = json.loads(path.read_text())
            except Exception:
                continue
            if not isinstance(case, dict) or not (case.get("plaintiff") or {}).get("name"):
                continue
            if target_org is not None and case_org(case) != target_org:
                continue
            key = (
                str(case.get("internal_case_number") or "").strip(),
                str((case.get("plaintiff") or {}).get("name") or "").strip().lower(),
            )
            prev = by_key.get(key)
            if not prev or str(case.get("captured_at") or "") > str((prev[1] or {}).get("captured_at") or ""):
                by_key[key] = (path, case)
    return list(by_key.values())


def save_case(path: Path, case: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(case, indent=2))


def capture_case_record(case: dict, org: str | None = None) -> str:
    """Assign an internal case number and write the record to storage."""
    pname = (case.get("plaintiff") or {}).get("name", "")
    initials = "".join(
        w[0].upper() for w in re.split(r"\s+", (pname or "").strip())
        if w and w[0].isalpha()
    )
    num = case.get("internal_case_number") or f"{datetime.now():%Y%m%d}-{initials or 'XX'}"
    case["internal_case_number"] = num
    case["organization"] = normalize_org(org or case_org(case))
    record = {**case, "captured_at": datetime.now().isoformat(timespec="seconds")}
    out_dir = primary_cases_dir()
    out_dir.mkdir(parents=True, exist_ok=True)
    save_case(out_dir / f"{num}_{slug(pname)}.json", record)
    return num


def import_case_json(raw: bytes, source_name: str = "upload", org: str | None = None) -> tuple[bool, str]:
    """Import a case JSON blob into the primary storage folder."""
    try:
        case = json.loads(raw.decode("utf-8"))
    except Exception as exc:
        return False, f"{source_name}: invalid JSON ({exc})"

    if not isinstance(case, dict):
        return False, f"{source_name}: top-level JSON must be an object"

    name = str((case.get("plaintiff") or {}).get("name") or "").strip()
    if not name:
        return False, f"{source_name}: missing plaintiff.name"

    case["organization"] = normalize_org(org or case_org(case))

    if not str(case.get("internal_case_number") or "").strip():
        capture_case_record(case)
    else:
        case["captured_at"] = case.get("captured_at") or datetime.now().isoformat(timespec="seconds")
        out_dir = primary_cases_dir()
        out_dir.mkdir(parents=True, exist_ok=True)
        save_case(out_dir / f"{case['internal_case_number']}_{slug(name)}.json", case)
    return True, f"Imported {name}"
