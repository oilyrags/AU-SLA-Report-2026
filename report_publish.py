#!/usr/bin/env python3
from __future__ import annotations

from typing import Dict, Optional

import pandas as pd

from sla_rules import PRIORITIES, get_targets


def _percent(numer: int, denom: int) -> Optional[float]:
    if denom == 0:
        return None
    return numer / denom


def build_summary_block(df: pd.DataFrame, quarter_label: str) -> pd.DataFrame:
    rows = []

    for priority in PRIORITIES:
        d = df[df["priority"] == priority].copy()
        tickets = len(d)

        frt_population = d[d["frt_has_sla"]]
        frt_met = int(((frt_population["frt_breached"] == False) & (frt_population["frt_pending"] == False)).sum())
        frt_breached = int((frt_population["frt_breached"] == True).sum())
        frt_pending = int((d["frt_pending"] == True).sum())
        frt_pct = _percent(frt_met, frt_met + frt_breached)

        res_population = d[d["res_has_sla"]]
        res_met = int(((res_population["res_breached"] == False) & (res_population["res_pending"] == False)).sum())
        res_breached = int((res_population["res_breached"] == True).sum())
        res_pending = int((d["res_pending"] == True).sum())
        res_pct = _percent(res_met, res_met + res_breached)

        if quarter_label == "YTD":
            frt_targets = [get_targets(priority, q)[0] for q in d["quarter"].dropna().tolist()]
            res_targets = [get_targets(priority, q)[1] for q in d["quarter"].dropna().tolist()]
            frt_target = sum(frt_targets) / len(frt_targets) if frt_targets else None
            res_target = sum(res_targets) / len(res_targets) if res_targets else None
        else:
            frt_target, res_target = get_targets(priority, quarter_label)

        rows.append(
            {
                "Quarter": quarter_label,
                "Priority": priority,
                "Tickets": tickets,
                "FRT Met": frt_met,
                "FRT Breached": frt_breached,
                "FRT Pending": frt_pending,
                "FRT %": frt_pct,
                "FRT Target": frt_target,
                "FRT Delta": (frt_pct - frt_target) if frt_pct is not None and frt_target is not None else None,
                "Res Met": res_met,
                "Res Breached": res_breached,
                "Res Pending": res_pending,
                "Res %": res_pct,
                "Res Target": res_target,
                "Res Delta": (res_pct - res_target) if res_pct is not None and res_target is not None else None,
            }
        )
    return pd.DataFrame(rows)


def build_dashboard_payload(in_scope: pd.DataFrame) -> pd.DataFrame:
    frt_population = in_scope[(in_scope["frt_has_sla"] == True) & (in_scope["frt_pending"] == False)]
    res_population = in_scope[(in_scope["res_has_sla"] == True) & (in_scope["res_pending"] == False)]

    frt_attainment = _percent(
        int((frt_population["frt_breached"] == False).sum()),
        len(frt_population),
    ) or 0.0
    res_attainment = _percent(
        int((res_population["res_breached"] == False).sum()),
        len(res_population),
    ) or 0.0

    open_risk = int(((in_scope["res_pending"] == True) & (in_scope["res_breached"] == True)).sum())

    return pd.DataFrame(
        [
            {"metric": "FRT Attainment %", "value": frt_attainment},
            {"metric": "Resolution Attainment %", "value": res_attainment},
            {"metric": "Tickets In Scope", "value": len(in_scope)},
            {"metric": "Open Risk", "value": open_risk},
        ]
    )


def build_refresh_control_payload(
    *,
    status: str,
    generated_at_iso: str,
    backup_reference: str,
    message: str,
    source_base_url: Optional[str] = None,
    request_metadata: Optional[Dict[str, str]] = None,
) -> pd.DataFrame:
    rows = [
        {"field": "last_refresh_status", "value": status},
        {"field": "generated_at", "value": generated_at_iso},
        {"field": "backup_reference", "value": backup_reference},
        {"field": "message", "value": message},
    ]
    if source_base_url is not None:
        rows.append({"field": "source_base_url", "value": source_base_url})
    for key in ("requested_at", "requested_by"):
        if request_metadata and request_metadata.get(key):
            rows.append({"field": key, "value": request_metadata[key]})
    return pd.DataFrame(rows)


def build_sheet_payloads(
    *,
    all_df: pd.DataFrame,
    in_scope: pd.DataFrame,
    exceptions_non_type: pd.DataFrame,
    exceptions_rejected: pd.DataFrame,
    exceptions_feature: pd.DataFrame,
    cycles_df: pd.DataFrame,
    base_url: str,
    generated_at_iso: str,
    backup_reference: str,
    request_metadata: Optional[Dict[str, str]] = None,
) -> Dict[str, pd.DataFrame]:
    refresh_control = build_refresh_control_payload(
        status="Succeeded",
        generated_at_iso=generated_at_iso,
        backup_reference=backup_reference,
        message="Refresh completed successfully.",
        source_base_url=base_url,
        request_metadata=request_metadata,
    )

    exceptions = pd.concat(
        [exceptions_non_type, exceptions_rejected, exceptions_feature],
        ignore_index=True,
    ).drop_duplicates(subset=["key"])

    return {
        "Dashboard": build_dashboard_payload(in_scope),
        "Summary": build_summary_block(in_scope, "YTD"),
        "Teams": in_scope[["team", "priority", "key"]].copy(),
        "Breaches": in_scope[(in_scope["frt_breached"] == True) | (in_scope["res_breached"] == True)].copy(),
        "Trends": in_scope[["week_start", "priority", "key"]].copy(),
        "Exceptions": exceptions,
        "Methodology": pd.DataFrame([{"field": "base_url", "value": base_url}]),
        "Raw Data": all_df.copy(),
        "Refresh Control": refresh_control,
    }


def publish_sheet_payloads(adapter, payloads: Dict[str, pd.DataFrame]) -> None:
    for tab_name, dataframe in payloads.items():
        adapter.write_tab(tab_name, dataframe)
