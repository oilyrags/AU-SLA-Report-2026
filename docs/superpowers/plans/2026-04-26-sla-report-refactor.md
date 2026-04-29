# SLA Report Refactor Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Refactor the SLA report codebase so rules are shared, responsibilities are smaller, refresh writes are safer, credentials are hardened, and developer setup is clearer.

**Architecture:** Add shared rule/config, Jira client, and workbook writer modules, then turn `sla_report_end2end.py` into a compatibility facade. Keep behavior stable by moving code first, then adding focused behavior changes with tests.

**Tech Stack:** Python 3.10, pandas, requests, openpyxl, Google Sheets REST API, Google Apps Script, bash, unittest.

---

## File Structure

- Create `sla_rules.py`: shared constants, report scope parsing, and pure helpers used by workbook and publish code.
- Create `jira_client.py`: Jira config, retrying API client, and prompt helper.
- Create `workbook_writer.py`: issue collection, scope classification, narrative generation, workbook creation, and worksheet helpers.
- Modify `sla_report_end2end.py`: compatibility re-exports and CLI entry point.
- Modify `report_publish.py`: remove duplicated targets and use `sla_rules.py`.
- Modify `report_refresh.py`: parameterized scope, atomic writes, request metadata preservation.
- Modify `google_sheets_adapter.py`: file permission hardening for token cache and private-key temp files.
- Modify `apps_script/README.md`: mention preserved request metadata and test command.
- Create `README.md`: setup, test, refresh, configuration, generated artifacts.
- Create `scripts/test.sh`: run the intended unittest command with `.venv`.
- Modify tests and add `test_sla_rules.py`: lock behavior during the refactor.

## Task 1: Shared Rules Module

**Files:**
- Create: `sla_rules.py`
- Test: `test_sla_rules.py`
- Modify: `report_publish.py`

- [ ] **Step 1: Write tests for shared scope and targets**

Create `test_sla_rules.py` with tests that assert:

```python
import os
import unittest
from unittest import mock

import sla_rules


class SlaRulesTests(unittest.TestCase):
    def test_report_scope_uses_defaults_env_and_explicit_values(self):
        with mock.patch.dict(os.environ, {}, clear=True):
            self.assertEqual(
                sla_rules.resolve_report_scope(project_key=None, year=None),
                ("ASD", 2026),
            )

        with mock.patch.dict(os.environ, {"JIRA_PROJECT_KEY": "OPS", "SLA_REPORT_YEAR": "2027"}, clear=True):
            self.assertEqual(
                sla_rules.resolve_report_scope(project_key=None, year=None),
                ("OPS", 2027),
            )
            self.assertEqual(
                sla_rules.resolve_report_scope(project_key="ASD", year=2026),
                ("ASD", 2026),
            )

    def test_report_scope_rejects_invalid_year(self):
        with self.assertRaisesRegex(ValueError, "SLA_REPORT_YEAR"):
            sla_rules.resolve_report_scope(project_key=None, year="twenty")

    def test_targets_and_aliases_are_shared(self):
        self.assertEqual(sla_rules.normalize_issue_type(" emal   request "), "email request")
        self.assertEqual(sla_rules.get_targets("P2 - High", "Q2"), (1.0, 0.95))
        self.assertEqual(sla_rules.get_targets("P2 – High", "Q2"), (1.0, 0.95))


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run the focused test and confirm it fails**

Run: `.venv/bin/python -m unittest test_sla_rules`

Expected: fail with `ModuleNotFoundError: No module named 'sla_rules'`.

- [ ] **Step 3: Implement `sla_rules.py`**

Move constants and helpers from `sla_report_end2end.py` into `sla_rules.py`, including `DEFAULT_PROJECT_KEY`, `DEFAULT_YEAR`, `PRIORITIES`, `TARGETS`, `IN_SCOPE_ISSUE_TYPES`, manual key sets, cue sets, `safe_get`, `extract_name`, `normalize_issue_type`, `parse_jira_dt`, `quarter_of`, `created_week_start`, `canonical_priority`, `hours_from_millis`, `percent`, `pctl`, `get_targets`, and `resolve_report_scope`.

- [ ] **Step 4: Update `report_publish.py` to import shared rules**

Remove local `PRIORITIES`, `TARGETS`, and `_get_targets`. Import `PRIORITIES` and `get_targets` from `sla_rules`, then replace `_get_targets(...)` calls with `get_targets(...)`.

- [ ] **Step 5: Run focused tests**

Run: `.venv/bin/python -m unittest test_sla_rules test_report_publish`

Expected: all selected tests pass.

## Task 2: Jira Client Module and Compatibility Facade

**Files:**
- Create: `jira_client.py`
- Modify: `sla_report_end2end.py`
- Modify: `workbook_writer.py`
- Test: existing suite

- [ ] **Step 1: Extract `jira_client.py`**

Move `JiraConfig`, `JiraClient`, `prompt_env`, retry constants, and Jira auth imports from `sla_report_end2end.py` to `jira_client.py`.

- [ ] **Step 2: Create `workbook_writer.py` from workbook/report logic**

Move `collect_issue_data`, `classify_scope`, `narrative_from_data`, `write_dataframe`, `build_workbook`, worksheet styling helpers, and workbook constants from `sla_report_end2end.py` to `workbook_writer.py`. Import shared helpers and constants from `sla_rules.py`.

- [ ] **Step 3: Replace `sla_report_end2end.py` with a facade**

Keep the executable script and re-export names used by tests:

```python
#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import os
from pathlib import Path

from jira_client import JiraClient, JiraConfig, prompt_env
from report_publish import build_dashboard_payload, build_sheet_payloads, build_summary_block
from report_refresh import build_refresh_status_payload, create_backup_bundle, run_local_refresh
from sla_rules import *
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
    print("[info] Checking Jira connectivity...", file=os.sys.stderr)
    client.get_server_info()

    print("[info] Pulling issues, true SLA cycles, and changelogs...", file=os.sys.stderr)
    all_df, cycles_df, raw_json = collect_issue_data(client, args.sleep_seconds, project_key=project_key, year=year)

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
```

- [ ] **Step 4: Run compatibility tests**

Run: `.venv/bin/python -m unittest test_sla_report_end2end test_report_publish`

Expected: all selected tests pass.

## Task 3: Parameterize Refresh and Preserve Request Metadata

**Files:**
- Modify: `report_refresh.py`
- Modify: `report_publish.py`
- Modify: `test_report_refresh.py`
- Modify: `test_report_publish.py`

- [ ] **Step 1: Add tests for refresh-control metadata**

Extend tests to cover `requested_at` and `requested_by` passing through `build_refresh_control_payload` and failure status payload creation.

- [ ] **Step 2: Update payload helpers**

Change `build_refresh_control_payload` to accept optional `request_metadata: dict | None`, append `requested_at` and `requested_by` rows when present, and keep existing field/value shape.

- [ ] **Step 3: Add adapter read support for request metadata**

Add an optional helper in `report_refresh.py`:

```python
def get_refresh_request_metadata(adapter) -> dict:
    if adapter is None or not hasattr(adapter, "read_tab_values"):
        return {}
    rows = adapter.read_tab_values("Refresh Control")
    values = {row[0]: row[1] for row in rows[1:] if len(row) >= 2}
    return {
        key: values[key]
        for key in ("requested_at", "requested_by")
        if values.get(key)
    }
```

- [ ] **Step 4: Use report scope in `report_refresh.py`**

Add `--project-key` and `--year` CLI args. Resolve them with `resolve_report_scope`. Pass `project_key` and `year` into `collect_issue_data` and `build_workbook`.

- [ ] **Step 5: Run refresh and publish tests**

Run: `.venv/bin/python -m unittest test_report_refresh test_report_publish`

Expected: all selected tests pass.

## Task 4: Safer Local Writes

**Files:**
- Modify: `report_refresh.py`
- Test: `test_report_refresh.py`

- [ ] **Step 1: Add tests for atomic raw JSON writes**

Add a test that monkeypatches `Path.replace` or validates the final raw JSON path exists and no sibling temp file remains after a successful refresh.

- [ ] **Step 2: Implement atomic byte writer**

Add:

```python
def write_bytes_atomically(path: Path, data: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = path.with_name(f".{path.name}.tmp")
    tmp_path.write_bytes(data)
    tmp_path.replace(path)
```

Use it for raw JSON. For workbook generation, write to a temporary workbook path, then replace the requested output path after `build_workbook` succeeds.

- [ ] **Step 3: Run refresh tests**

Run: `.venv/bin/python -m unittest test_report_refresh`

Expected: all selected tests pass.

## Task 5: Credential File Hardening

**Files:**
- Modify: `google_sheets_adapter.py`
- Modify: `test_google_sheets_adapter.py`

- [ ] **Step 1: Add token-cache mode test**

Add a test that writes a token through `_write_cached_token`, then checks `stat.S_IMODE(token_path.stat().st_mode) == 0o600`.

- [ ] **Step 2: Add private-key temp mode test**

Patch `subprocess.run` and `os.unlink`, call `_build_jwt_assertion`, capture the key path passed to `openssl`, and assert the file mode is `0o600` before unlink cleanup.

- [ ] **Step 3: Implement permission hardening**

Use `os.open(..., 0o600)` for token and private-key writes, wrap the descriptor with `open(fd, "w", encoding="utf-8")`, and keep existing cleanup behavior.

- [ ] **Step 4: Run Google Sheets tests**

Run: `.venv/bin/python -m unittest test_google_sheets_adapter`

Expected: all selected tests pass.

## Task 6: Developer Entry Points and Docs

**Files:**
- Create: `README.md`
- Create: `scripts/test.sh`
- Modify: `apps_script/README.md`

- [ ] **Step 1: Add `scripts/test.sh`**

Create an executable bash script that verifies `.venv/bin/python` exists and runs `.venv/bin/python -m unittest "$@"`.

- [ ] **Step 2: Add root README**

Document setup, dependencies, test command, refresh command, environment variables, generated artifacts, and Google Sheets publish behavior.

- [ ] **Step 3: Update Apps Script README**

Mention that Python preserves `requested_at` and `requested_by` when the adapter can read existing refresh-control values.

- [ ] **Step 4: Run doc-independent test command**

Run: `./scripts/test.sh`

Expected: all tests pass.

## Task 7: Final Cleanup and Verification

**Files:**
- Review all changed Python and docs files.

- [ ] **Step 1: Search for duplicated target tables and stale imports**

Run: `rg -n "TARGETS =|from collections import Counter|TODO|TBD|implement later" .`

Expected: one target table in `sla_rules.py`, no stale `Counter`, no new placeholders.

- [ ] **Step 2: Run full test suite**

Run: `./scripts/test.sh`

Expected: all tests pass.

- [ ] **Step 3: Summarize changed files and residual risks**

Report the module split, behavior fixes, and verification result. Note that no live Jira or Google API call was performed.
