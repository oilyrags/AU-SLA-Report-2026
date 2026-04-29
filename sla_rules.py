#!/usr/bin/env python3
from __future__ import annotations

import math
import os
import re
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple


SEARCH_PAGE_SIZE = 100
RETRY_STATUS_CODES = {429, 500, 502, 503, 504}
DEFAULT_TIMEOUT = 60

DEFAULT_PROJECT_KEY = "ASD"
DEFAULT_YEAR = 2026

BUG_TYPE = "Report a bug"
ISSUE_TYPE = "Technical support"
EMAIL_REQUEST_TYPE = "Email request"
IN_SCOPE_ISSUE_TYPES = {BUG_TYPE, ISSUE_TYPE, EMAIL_REQUEST_TYPE}
IN_SCOPE_ISSUE_TYPE_ALIASES = {
    "emal request": EMAIL_REQUEST_TYPE.lower(),
}

NOT_ACCEPTED_RESOLUTIONS = {"Rejected", "Won't Do", "Duplicate", "Declined"}

MANUAL_INCLUDE_BUG_KEYS = {"ASD-1999", "ASD-2000"}
MANUAL_DISGUISED_FEATURE_KEYS = {"ASD-1992", "ASD-1993"}

FEATURE_REQUEST_CUES = (
    "i want to",
    "please",
    "request",
    "new feature",
    "feature",
    "bulk edit",
    "searchfunction",
    "acceptance criteria",
    "create",
    "add",
    "implement",
)

BUG_CUES = (
    "error",
    "does not work",
    "doesn't work",
    "not working",
    "fails",
    "failed",
    "broken",
    "cannot",
    "can't",
    "unable",
    "incident",
)

TEAM_FIELD_NAME = "Domain"
TEAM_FIELD_ID = "customfield_11244"

SLA_FRT_NAMES = {"time to first response"}
SLA_RES_NAMES = {"time to resolution", "time to done", "time to resolution after agreement"}

Q1_TARGET_MODE = "Q2"

PRIORITIES = ["P1 – Critical", "P2 – High", "P3 – Medium", "P4 – Low"]

PRIORITY_CANON = {
    "Critical": "P1 – Critical",
    "High": "P2 – High",
    "Medium": "P3 – Medium",
    "Low": "P4 – Low",
    "P1 - Critical": "P1 – Critical",
    "P2 - High": "P2 – High",
    "P3 - Medium": "P3 – Medium",
    "P4 - Low": "P4 – Low",
}

TARGETS = {
    "P1 – Critical": {
        "Q1": {"frt_pct": 1.00, "res_pct": 1.00},
        "Q2": {"frt_pct": 1.00, "res_pct": 1.00},
        "Q3": {"frt_pct": 1.00, "res_pct": 1.00},
        "Q4": {"frt_pct": 1.00, "res_pct": 1.00},
    },
    "P2 – High": {
        "Q1": {"frt_pct": 1.00, "res_pct": 0.95},
        "Q2": {"frt_pct": 1.00, "res_pct": 0.95},
        "Q3": {"frt_pct": 1.00, "res_pct": 0.95},
        "Q4": {"frt_pct": 1.00, "res_pct": 0.95},
    },
    "P3 – Medium": {
        "Q1": {"frt_pct": 0.90, "res_pct": 0.90},
        "Q2": {"frt_pct": 0.90, "res_pct": 0.90},
        "Q3": {"frt_pct": 0.95, "res_pct": 0.95},
        "Q4": {"frt_pct": 0.95, "res_pct": 0.95},
    },
    "P4 – Low": {
        "Q1": {"frt_pct": 0.90, "res_pct": 0.90},
        "Q2": {"frt_pct": 0.90, "res_pct": 0.90},
        "Q3": {"frt_pct": 0.95, "res_pct": 0.95},
        "Q4": {"frt_pct": 0.95, "res_pct": 0.95},
    },
}


def resolve_report_scope(project_key: Optional[str], year: Optional[int | str]) -> Tuple[str, int]:
    resolved_project_key = (project_key or os.environ.get("JIRA_PROJECT_KEY") or DEFAULT_PROJECT_KEY).strip()
    if not resolved_project_key:
        raise ValueError("JIRA_PROJECT_KEY cannot be blank")

    raw_year = year if year is not None else os.environ.get("SLA_REPORT_YEAR", DEFAULT_YEAR)
    try:
        resolved_year = int(raw_year)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"SLA_REPORT_YEAR must be an integer, got {raw_year!r}") from exc

    return resolved_project_key, resolved_year


def safe_get(d: Any, *keys: str) -> Any:
    cur = d
    for key in keys:
        if not isinstance(cur, dict):
            return None
        cur = cur.get(key)
    return cur


def extract_name(obj: Any) -> Optional[str]:
    if obj is None:
        return None
    if isinstance(obj, dict):
        return obj.get("displayName") or obj.get("name") or obj.get("value")
    if isinstance(obj, list):
        vals = [extract_name(x) for x in obj]
        vals = [v for v in vals if v]
        return ", ".join(vals) if vals else None
    return str(obj)


def normalize_issue_type(value: Optional[str]) -> Optional[str]:
    if not value:
        return None
    normalized = re.sub(r"\s+", " ", str(value).strip()).lower()
    return IN_SCOPE_ISSUE_TYPE_ALIASES.get(normalized, normalized)


def parse_jira_dt(value: Optional[str]) -> Optional[datetime]:
    if not value:
        return None
    fmts = ["%Y-%m-%dT%H:%M:%S.%f%z", "%Y-%m-%dT%H:%M:%S%z"]
    for fmt in fmts:
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            pass
    return None


def quarter_of(created_str: Optional[str]) -> Optional[str]:
    dt = parse_jira_dt(created_str)
    if not dt:
        return None
    return f"Q{((dt.month - 1) // 3) + 1}"


def created_week_start(created_str: Optional[str]) -> Optional[str]:
    dt = parse_jira_dt(created_str)
    if not dt:
        return None
    monday = dt.date().fromordinal(dt.date().toordinal() - dt.weekday())
    return monday.isoformat()


def canonical_priority(priority_name: Optional[str]) -> Optional[str]:
    if not priority_name:
        return None
    return PRIORITY_CANON.get(priority_name, priority_name)


def hours_from_millis(ms: Any) -> Optional[float]:
    if ms is None:
        return None
    try:
        return float(ms) / 1000.0 / 60.0 / 60.0
    except Exception:
        return None


def percent(numer: int, denom: int) -> Optional[float]:
    if denom == 0:
        return None
    return numer / denom


def pctl(values: List[float], p: float) -> Optional[float]:
    if not values:
        return None
    s = sorted(values)
    if len(s) == 1:
        return s[0]
    idx = (len(s) - 1) * p
    lo = math.floor(idx)
    hi = math.ceil(idx)
    if lo == hi:
        return s[int(idx)]
    frac = idx - lo
    return s[lo] * (1 - frac) + s[hi] * frac


def get_targets(priority: Optional[str], quarter: Optional[str]) -> Tuple[Optional[float], Optional[float]]:
    if not priority or not quarter:
        return None, None
    canonical = canonical_priority(priority)
    tgt = TARGETS.get(canonical, {}).get(quarter)
    if not tgt:
        return None, None
    return tgt["frt_pct"], tgt["res_pct"]
