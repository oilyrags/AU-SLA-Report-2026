#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import os
import sys
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd

from report_publish import build_refresh_control_payload, build_sheet_payloads, publish_sheet_payloads
from sla_rules import resolve_report_scope


def build_backup_reference(generated_at_iso: str) -> str:
    generated_at = datetime.fromisoformat(generated_at_iso.replace("Z", "+00:00"))
    return generated_at.strftime("backup-%Y%m%d-%H%M%S-%f")


def create_backup_bundle(
    *,
    backup_root: Path,
    raw_json_bytes: bytes,
    workbook_bytes: bytes,
    published_payload_bytes: bytes,
    generated_at_iso: str,
):
    backup_reference = build_backup_reference(generated_at_iso)
    backup_dir = backup_root / backup_reference
    backup_dir.mkdir(parents=True, exist_ok=False)

    (backup_dir / "raw_jira_extract.json").write_bytes(raw_json_bytes)
    (backup_dir / "report.xlsx").write_bytes(workbook_bytes)
    (backup_dir / "sheet_payloads.json").write_bytes(published_payload_bytes)

    return {"backup_reference": backup_reference, "backup_dir": str(backup_dir)}


def temporary_sibling_path(path: Path) -> Path:
    return path.with_name(f".{path.name}.tmp")


def write_bytes_atomically(path: Path, data: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = temporary_sibling_path(path)
    tmp_path.write_bytes(data)
    tmp_path.replace(path)


def build_refresh_status_payload(
    *,
    status: str,
    generated_at_iso: str,
    backup_reference: str,
    message: str,
    request_metadata: dict | None = None,
) -> pd.DataFrame:
    return build_refresh_control_payload(
        status=status,
        generated_at_iso=generated_at_iso,
        backup_reference=backup_reference,
        message=message,
        request_metadata=request_metadata,
    )


def payloads_to_json_bytes(payloads) -> bytes:
    from google_sheets_adapter import dataframe_to_values

    serializable = {}
    for tab_name, payload in payloads.items():
        if isinstance(payload, pd.DataFrame):
            serializable[tab_name] = dataframe_to_values(payload)
        else:
            serializable[tab_name] = payload
    return json.dumps(serializable, ensure_ascii=False, indent=2, default=str).encode("utf-8")


def collect_issue_data(client, sleep_seconds, **kwargs):
    from workbook_writer import collect_issue_data as _collect_issue_data

    return _collect_issue_data(client, sleep_seconds, **kwargs)


def classify_scope(df):
    from workbook_writer import classify_scope as _classify_scope

    return _classify_scope(df)


def build_workbook(**kwargs):
    from workbook_writer import build_workbook as _build_workbook

    return _build_workbook(**kwargs)


def build_google_sheets_adapter_from_env():
    from google_sheets_adapter import load_google_sheets_adapter_from_env

    return load_google_sheets_adapter_from_env()


def build_refresh_payloads_from_frames(
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
    request_metadata: dict | None = None,
):
    return build_sheet_payloads(
        all_df=all_df,
        in_scope=in_scope,
        exceptions_non_type=exceptions_non_type,
        exceptions_rejected=exceptions_rejected,
        exceptions_feature=exceptions_feature,
        cycles_df=cycles_df,
        base_url=base_url,
        generated_at_iso=generated_at_iso,
        backup_reference=backup_reference,
        request_metadata=request_metadata,
    )


def get_refresh_request_metadata(adapter) -> dict:
    if adapter is None or not hasattr(adapter, "read_tab_values"):
        return {}
    try:
        rows = adapter.read_tab_values("Refresh Control")
    except Exception as exc:
        print(f"[warn] Failed to read existing refresh request metadata: {exc}", file=sys.stderr)
        return {}

    values = {}
    for row in rows[1:]:
        if len(row) >= 2:
            values[str(row[0])] = row[1]

    return {
        key: values[key]
        for key in ("requested_at", "requested_by")
        if values.get(key)
    }


def run_local_refresh(
    *,
    client,
    adapter,
    generated_at_iso: str,
    base_url: str,
    project_key: str | None = None,
    year: int | None = None,
):
    collect_kwargs = {}
    if project_key is not None:
        collect_kwargs["project_key"] = project_key
    if year is not None:
        collect_kwargs["year"] = year
    all_df, cycles_df, _raw_json = collect_issue_data(client, sleep_seconds=0.0, **collect_kwargs)
    all_df, in_scope, exceptions_non_type, exceptions_rejected, exceptions_feature = classify_scope(all_df)
    request_metadata = get_refresh_request_metadata(adapter)
    payloads = build_refresh_payloads_from_frames(
        all_df=all_df,
        in_scope=in_scope,
        exceptions_non_type=exceptions_non_type,
        exceptions_rejected=exceptions_rejected,
        exceptions_feature=exceptions_feature,
        cycles_df=cycles_df,
        base_url=base_url,
        generated_at_iso=generated_at_iso,
        backup_reference="pending-backup",
        request_metadata=request_metadata,
    )
    publish_sheet_payloads(adapter, payloads)
    return payloads


def refresh_report_local(
    *,
    client,
    base_url: str,
    generated_at_iso: str,
    output_path: Path,
    raw_json_path: Path,
    backup_root: Path | None = None,
    sleep_seconds: float = 0.15,
    google_adapter=None,
    project_key: str | None = None,
    year: int | None = None,
):
    collect_kwargs = {}
    workbook_kwargs = {}
    if project_key is not None:
        collect_kwargs["project_key"] = project_key
        workbook_kwargs["project_key"] = project_key
    if year is not None:
        collect_kwargs["year"] = year
        workbook_kwargs["year"] = year

    all_df, cycles_df, raw_json = collect_issue_data(client, sleep_seconds=sleep_seconds, **collect_kwargs)
    all_df, in_scope, exceptions_non_type, exceptions_rejected, exceptions_feature = classify_scope(all_df)

    raw_json_bytes = json.dumps(raw_json, ensure_ascii=False, indent=2).encode("utf-8")
    write_bytes_atomically(raw_json_path, raw_json_bytes)

    temp_output_path = temporary_sibling_path(output_path)
    counts = build_workbook(
        all_df=all_df,
        in_scope=in_scope,
        exceptions_non_type=exceptions_non_type,
        exceptions_rejected=exceptions_rejected,
        exceptions_feature=exceptions_feature,
        cycles_df=cycles_df,
        raw_json_path=str(raw_json_path),
        output_path=str(temp_output_path),
        base_url=base_url,
        **workbook_kwargs,
    )
    temp_output_path.replace(output_path)

    backup_reference = build_backup_reference(generated_at_iso)
    adapter = google_adapter if google_adapter is not None else build_google_sheets_adapter_from_env()
    request_metadata = get_refresh_request_metadata(adapter)
    payloads = build_refresh_payloads_from_frames(
        all_df=all_df,
        in_scope=in_scope,
        exceptions_non_type=exceptions_non_type,
        exceptions_rejected=exceptions_rejected,
        exceptions_feature=exceptions_feature,
        cycles_df=cycles_df,
        base_url=base_url,
        generated_at_iso=generated_at_iso,
        backup_reference=backup_reference,
        request_metadata=request_metadata,
    )
    published_payload_bytes = payloads_to_json_bytes(payloads)
    backup = create_backup_bundle(
        backup_root=backup_root or output_path.parent / "backups",
        raw_json_bytes=raw_json_bytes,
        workbook_bytes=output_path.read_bytes(),
        published_payload_bytes=published_payload_bytes,
        generated_at_iso=generated_at_iso,
    )

    google_published = False
    if adapter is not None:
        try:
            publish_sheet_payloads(adapter, payloads)
            google_published = True
        except Exception as exc:
            failure_payloads = {
                "Refresh Control": build_refresh_status_payload(
                    status="Failed",
                    generated_at_iso=generated_at_iso,
                    backup_reference=backup_reference,
                    message=f"Refresh failed: {exc}",
                    request_metadata=request_metadata,
                )
            }
            try:
                publish_sheet_payloads(adapter, failure_payloads)
            except Exception as status_exc:
                print(f"[warn] Failed to publish refresh failure status: {status_exc}", file=sys.stderr)
            raise
    else:
        print("[info] Google Sheets config not found; skipping Google Sheets publish.", file=sys.stderr)

    return {
        "counts": counts,
        "payloads": payloads,
        "google_published": google_published,
        "raw_json_path": str(raw_json_path),
        "output_path": str(output_path),
        "backup_reference": backup["backup_reference"],
        "backup_dir": backup["backup_dir"],
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--sleep-seconds", type=float, default=0.15, help="Delay between per-issue SLA/changelog calls")
    parser.add_argument("--project-key", default=None, help="Jira project key. Defaults to JIRA_PROJECT_KEY or ASD.")
    parser.add_argument("--year", default=None, help="Report year. Defaults to SLA_REPORT_YEAR or 2026.")
    args = parser.parse_args()

    from jira_client import JiraClient, JiraConfig, prompt_env

    project_key, year = resolve_report_scope(project_key=args.project_key, year=args.year)
    base_url = prompt_env("JIRA_BASE_URL", "Jira base URL")
    email = prompt_env("JIRA_EMAIL", "Jira email")
    token = prompt_env("JIRA_API_TOKEN", "Jira API token", secret=True)

    cfg = JiraConfig(base_url=base_url, email=email, api_token=token)
    client = JiraClient(cfg)

    print("[info] Checking Jira connectivity...", file=sys.stderr)
    client.get_server_info()

    print("[info] Pulling issues, true SLA cycles, and changelogs...", file=sys.stderr)
    generated_at_iso = datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")
    result = refresh_report_local(
        client=client,
        base_url=base_url,
        generated_at_iso=generated_at_iso,
        output_path=Path(os.path.abspath(f"SMG_Automotive_SLA_Report_{year}YTD.xlsx")),
        raw_json_path=Path(os.path.abspath(f"raw_jira_extract_{year}ytd.json")),
        sleep_seconds=args.sleep_seconds,
        project_key=project_key,
        year=year,
    )

    counts = result["counts"]
    print("")
    print("Done.")
    print(f"Workbook: {result['output_path']}")
    print(f"Raw JSON: {result['raw_json_path']}")
    print(
        f"Pulled {counts['pulled']} tickets. "
        f"{counts['in_scope']} in-scope. "
        f"{counts['exceptions_total']} in exceptions "
        f"({counts['exceptions_non_type']} non-Bug/Issue, "
        f"{counts['exceptions_rejected']} rejected/won't-do, "
        f"{counts['exceptions_feature']} feature requests/manual exclusions)."
    )
    if result["google_published"]:
        print("Google Sheets publish: completed.")


if __name__ == "__main__":
    main()
