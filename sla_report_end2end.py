#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from jira_client import JiraClient, JiraConfig, prompt_env
from report_publish import build_dashboard_payload, build_sheet_payloads, build_summary_block
from report_refresh import build_refresh_status_payload, create_backup_bundle, run_local_refresh
from sla_rules import *  # noqa: F403 - preserve the old script's import surface.
from sla_rules import resolve_report_scope
from workbook_writer import build_workbook, classify_scope, collect_issue_data, narrative_from_data, write_dataframe


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--sleep-seconds", type=float, default=0.15, help="Delay between per-issue SLA/changelog calls")
    parser.add_argument("--project-key", default=None, help="Jira project key. Defaults to JIRA_PROJECT_KEY or ASD.")
    parser.add_argument("--year", default=None, help="Report year. Defaults to SLA_REPORT_YEAR or 2026.")
    args = parser.parse_args()

    project_key, year = resolve_report_scope(project_key=args.project_key, year=args.year)
    base_url = prompt_env("JIRA_BASE_URL", "Jira base URL")
    email = prompt_env("JIRA_EMAIL", "Jira email")
    token = prompt_env("JIRA_API_TOKEN", "Jira API token", secret=True)

    client = JiraClient(JiraConfig(base_url=base_url, email=email, api_token=token))

    print("[info] Checking Jira connectivity...", file=sys.stderr)
    client.get_server_info()

    print("[info] Pulling issues, true SLA cycles, and changelogs...", file=sys.stderr)
    all_df, cycles_df, raw_json = collect_issue_data(
        client,
        args.sleep_seconds,
        project_key=project_key,
        year=year,
    )

    raw_json_path = Path.cwd() / f"raw_jira_extract_{year}ytd.json"
    raw_json_path.write_text(json.dumps(raw_json, ensure_ascii=False, indent=2), encoding="utf-8")

    all_df, in_scope, exceptions_non_type, exceptions_rejected, exceptions_feature = classify_scope(all_df)

    output_path = Path.cwd() / f"SMG_Automotive_SLA_Report_{year}YTD.xlsx"
    counts = build_workbook(
        all_df=all_df,
        in_scope=in_scope,
        exceptions_non_type=exceptions_non_type,
        exceptions_rejected=exceptions_rejected,
        exceptions_feature=exceptions_feature,
        cycles_df=cycles_df,
        raw_json_path=str(raw_json_path),
        output_path=str(output_path),
        base_url=base_url,
        project_key=project_key,
        year=year,
    )

    print("")
    print("Done.")
    print(f"Workbook: {output_path}")
    print(f"Raw JSON: {raw_json_path}")
    print(
        f"Pulled {counts['pulled']} tickets. "
        f"{counts['in_scope']} in-scope. "
        f"{counts['exceptions_total']} in exceptions "
        f"({counts['exceptions_non_type']} non-Bug/Issue, "
        f"{counts['exceptions_rejected']} rejected/won't-do, "
        f"{counts['exceptions_feature']} feature requests/manual exclusions)."
    )


if __name__ == "__main__":
    main()
