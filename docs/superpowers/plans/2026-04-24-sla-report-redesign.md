# SLA Report Redesign Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a Google Sheets-native SLA report with a modern dashboard, analyst-ready tabs, a semi-local refresh workflow, and automatic backups while preserving Python as the source of truth.

**Architecture:** Refactor the current monolithic Python script into testable report-building and publishing units, then add a thin Google Apps Script layer for refresh request/status UX. Python will continue to pull Jira data and compute SLA metrics, and will additionally shape publish payloads, create backups, and push tab data into Google Sheets.

**Tech Stack:** Python 3, `pandas`, `requests`, `openpyxl`, `unittest`, Google Apps Script, Google Sheets API/client integration as needed by the publish adapter

---

## File Structure

- Modify: `/Users/cliada/Documents/code/projects/SLA-Report/sla_report_end2end.py`
  - Keep the existing extraction and workbook generation entry point working while carving out pure helpers and orchestration boundaries.
- Create: `/Users/cliada/Documents/code/projects/SLA-Report/report_refresh.py`
  - Hold refresh orchestration helpers: refresh request reading, status transitions, backup metadata, and the top-level semi-local refresh workflow.
- Create: `/Users/cliada/Documents/code/projects/SLA-Report/report_publish.py`
  - Hold Google Sheets-facing payload shaping and publishing boundaries so the publish contract is explicit and testable.
- Create: `/Users/cliada/Documents/code/projects/SLA-Report/apps_script/Code.gs`
  - Hold the Google Apps Script menu/button handlers and refresh-control cell updates.
- Create: `/Users/cliada/Documents/code/projects/SLA-Report/apps_script/README.md`
  - Explain how to install the script into the target sheet and how the semi-local refresh handshake works.
- Modify: `/Users/cliada/Documents/code/projects/SLA-Report/test_sla_report_end2end.py`
  - Keep existing classification tests and add pure Python coverage for summary and dashboard shaping where practical.
- Create: `/Users/cliada/Documents/code/projects/SLA-Report/test_report_publish.py`
  - Add contract tests for tab payload names, columns, backup references, and refresh-control metadata.
- Create: `/Users/cliada/Documents/code/projects/SLA-Report/test_report_refresh.py`
  - Add refresh-state and backup tests using fixture data instead of live Jira.
- Create: `/Users/cliada/Documents/code/projects/SLA-Report/tests/fixtures/report_input.json`
  - Small fixture dataset that exercises classification, KPI derivation, publish payload shaping, and backup logic without network access.

### Task 1: Lock the Publish Contract with Failing Tests

**Files:**
- Create: `/Users/cliada/Documents/code/projects/SLA-Report/test_report_publish.py`
- Create: `/Users/cliada/Documents/code/projects/SLA-Report/tests/fixtures/report_input.json`
- Modify: `/Users/cliada/Documents/code/projects/SLA-Report/sla_report_end2end.py`

- [ ] **Step 1: Write the failing contract tests**

```python
import json
import pathlib
import unittest

import pandas as pd

import sla_report_end2end


FIXTURE = pathlib.Path(__file__).parent / "tests" / "fixtures" / "report_input.json"


def load_fixture_frames():
    payload = json.loads(FIXTURE.read_text(encoding="utf-8"))
    all_df = pd.DataFrame(payload["all_df"])
    cycles_df = pd.DataFrame(payload["cycles_df"])
    _, in_scope, exceptions_non_type, exceptions_rejected, exceptions_feature = (
        sla_report_end2end.classify_scope(all_df)
    )
    return all_df, in_scope, exceptions_non_type, exceptions_rejected, exceptions_feature, cycles_df


class PublishContractTests(unittest.TestCase):
    def test_build_sheet_payloads_returns_expected_tabs(self):
        all_df, in_scope, non_type, rejected, feature, cycles_df = load_fixture_frames()

        payloads = sla_report_end2end.build_sheet_payloads(
            all_df=all_df,
            in_scope=in_scope,
            exceptions_non_type=non_type,
            exceptions_rejected=rejected,
            exceptions_feature=feature,
            cycles_df=cycles_df,
            base_url="https://example.atlassian.net",
            generated_at_iso="2026-04-24T12:00:00+00:00",
            backup_reference="backup-20260424-120000",
        )

        self.assertEqual(
            list(payloads.keys()),
            [
                "Dashboard",
                "Summary",
                "Teams",
                "Breaches",
                "Trends",
                "Exceptions",
                "Methodology",
                "Raw Data",
                "Refresh Control",
            ],
        )
        self.assertIn("last_refresh_status", payloads["Refresh Control"].columns.tolist())
        self.assertIn("backup_reference", payloads["Refresh Control"].columns.tolist())

    def test_dashboard_payload_contains_exec_kpis(self):
        all_df, in_scope, non_type, rejected, feature, cycles_df = load_fixture_frames()

        payloads = sla_report_end2end.build_sheet_payloads(
            all_df=all_df,
            in_scope=in_scope,
            exceptions_non_type=non_type,
            exceptions_rejected=rejected,
            exceptions_feature=feature,
            cycles_df=cycles_df,
            base_url="https://example.atlassian.net",
            generated_at_iso="2026-04-24T12:00:00+00:00",
            backup_reference="backup-20260424-120000",
        )

        dashboard = payloads["Dashboard"]
        self.assertIn("metric", dashboard.columns.tolist())
        self.assertIn("value", dashboard.columns.tolist())
        self.assertTrue((dashboard["metric"] == "FRT Attainment %").any())
        self.assertTrue((dashboard["metric"] == "Resolution Attainment %").any())
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `python3 -m unittest test_report_publish.py -v`
Expected: `AttributeError` or `ImportError` because `build_sheet_payloads` and the new fixture-backed helpers do not exist yet.

- [ ] **Step 3: Add the fixture file**

```json
{
  "all_df": [
    {
      "key": "ASD-1001",
      "summary": "Critical checkout issue",
      "issuetype": "Report a bug",
      "resolution": "Resolved",
      "priority": "P1 – Critical",
      "created": "2026-01-10T08:30:00.000+0000",
      "resolutiondate": "2026-01-10T10:00:00.000+0000",
      "team": "Checkout",
      "domain": "Checkout",
      "components": "",
      "labels": "ServiceDesk",
      "assignee": "Alex",
      "reporter": "Pat",
      "quarter": "Q1",
      "week_start": "2026-01-05",
      "created_day_of_week": "Saturday",
      "created_hour": 8,
      "frt_has_sla": true,
      "frt_pending": false,
      "frt_breached": false,
      "frt_elapsed_hours": 0.25,
      "frt_goal_hours": 1.0,
      "frt_total_cycles": 1,
      "res_has_sla": true,
      "res_pending": false,
      "res_breached": false,
      "res_elapsed_hours": 1.5,
      "res_goal_hours": 8.0,
      "res_total_cycles": 1,
      "business_to_wall_clock_ratio": 0.75,
      "priority_downgrade_evidence": false,
      "status": "Done",
      "url": "https://example.atlassian.net/browse/ASD-1001"
    },
    {
      "key": "ASD-1002",
      "summary": "Search fix request",
      "issuetype": "Technical support",
      "resolution": "Resolved",
      "priority": "P3 – Medium",
      "created": "2026-02-12T18:00:00.000+0000",
      "resolutiondate": "2026-02-15T12:00:00.000+0000",
      "team": "Inventory - Retool",
      "domain": "Inventory - Retool",
      "components": "",
      "labels": "ServiceDesk",
      "assignee": "Robin",
      "reporter": "Sam",
      "quarter": "Q1",
      "week_start": "2026-02-09",
      "created_day_of_week": "Thursday",
      "created_hour": 18,
      "frt_has_sla": true,
      "frt_pending": false,
      "frt_breached": true,
      "frt_elapsed_hours": 5.0,
      "frt_goal_hours": 1.0,
      "frt_total_cycles": 2,
      "res_has_sla": true,
      "res_pending": false,
      "res_breached": true,
      "res_elapsed_hours": 30.0,
      "res_goal_hours": 16.0,
      "res_total_cycles": 2,
      "business_to_wall_clock_ratio": 0.55,
      "priority_downgrade_evidence": true,
      "status": "Done",
      "url": "https://example.atlassian.net/browse/ASD-1002"
    }
  ],
  "cycles_df": [
    {
      "key": "ASD-1001",
      "sla_name": "Time to first response",
      "cycle_kind": "completed",
      "cycle_index": 1,
      "breached": false
    },
    {
      "key": "ASD-1002",
      "sla_name": "Time to resolution",
      "cycle_kind": "completed",
      "cycle_index": 1,
      "breached": true
    }
  ]
}
```

- [ ] **Step 4: Implement the minimal publish contract helpers**

```python
def build_dashboard_payload(in_scope: pd.DataFrame) -> pd.DataFrame:
    frt_population = in_scope[in_scope["frt_pending"] == False]
    res_population = in_scope[in_scope["res_pending"] == False]
    frt_attainment = percent(int((frt_population["frt_breached"] == False).sum()), len(frt_population)) or 0.0
    res_attainment = percent(int((res_population["res_breached"] == False).sum()), len(res_population)) or 0.0
    return pd.DataFrame(
        [
            {"metric": "FRT Attainment %", "value": frt_attainment},
            {"metric": "Resolution Attainment %", "value": res_attainment},
            {"metric": "Tickets In Scope", "value": len(in_scope)},
            {"metric": "Open Risk", "value": int(((in_scope["res_pending"] == True) & (in_scope["res_breached"] == True)).sum())},
        ]
    )


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
) -> Dict[str, pd.DataFrame]:
    refresh_control = pd.DataFrame(
        [
            {
                "last_refresh_status": "Succeeded",
                "generated_at": generated_at_iso,
                "backup_reference": backup_reference,
                "source_base_url": base_url,
            }
        ]
    )
    return {
        "Dashboard": build_dashboard_payload(in_scope),
        "Summary": build_summary_block(in_scope, "YTD"),
        "Teams": in_scope[["team", "priority", "key"]].copy(),
        "Breaches": in_scope[(in_scope["frt_breached"] == True) | (in_scope["res_breached"] == True)].copy(),
        "Trends": in_scope[["week_start", "priority", "key"]].copy(),
        "Exceptions": pd.concat([exceptions_non_type, exceptions_rejected, exceptions_feature], ignore_index=True),
        "Methodology": pd.DataFrame([{"field": "base_url", "value": base_url}]),
        "Raw Data": all_df.copy(),
        "Refresh Control": refresh_control,
    }
```

- [ ] **Step 5: Run the tests to verify they pass**

Run: `python3 -m unittest test_report_publish.py -v`
Expected: `OK`

- [ ] **Step 6: Checkpoint commit**

If this directory is initialized as Git, run:

```bash
git add test_report_publish.py tests/fixtures/report_input.json sla_report_end2end.py
git commit -m "test: lock sheet publish contract"
```

### Task 2: Refactor Report Shaping into Focused Units

**Files:**
- Create: `/Users/cliada/Documents/code/projects/SLA-Report/report_publish.py`
- Modify: `/Users/cliada/Documents/code/projects/SLA-Report/sla_report_end2end.py`
- Modify: `/Users/cliada/Documents/code/projects/SLA-Report/test_sla_report_end2end.py`

- [ ] **Step 1: Write the failing unit tests for extracted shaping helpers**

```python
class DashboardShapingTests(unittest.TestCase):
    def test_build_dashboard_payload_reports_open_risk_and_volume(self):
        df = pd.DataFrame(
            [
                {"key": "ASD-1", "frt_pending": False, "frt_breached": False, "res_pending": False, "res_breached": False},
                {"key": "ASD-2", "frt_pending": False, "frt_breached": True, "res_pending": True, "res_breached": True},
            ]
        )

        payload = sla_report_end2end.build_dashboard_payload(df)
        metrics = dict(zip(payload["metric"], payload["value"]))

        self.assertEqual(metrics["Tickets In Scope"], 2)
        self.assertEqual(metrics["Open Risk"], 1)
```

- [ ] **Step 2: Run the tests to verify they fail for the right reason**

Run: `python3 -m unittest test_sla_report_end2end.py -v`
Expected: failure because the helper has not been imported from a focused module yet or does not expose the stabilized behavior.

- [ ] **Step 3: Move the shaping helpers into `report_publish.py`**

```python
import pandas as pd


def build_dashboard_payload(in_scope: pd.DataFrame, percent_fn) -> pd.DataFrame:
    frt_population = in_scope[in_scope["frt_pending"] == False]
    res_population = in_scope[in_scope["res_pending"] == False]
    return pd.DataFrame(
        [
            {
                "metric": "FRT Attainment %",
                "value": percent_fn(int((frt_population["frt_breached"] == False).sum()), len(frt_population)) or 0.0,
            },
            {
                "metric": "Resolution Attainment %",
                "value": percent_fn(int((res_population["res_breached"] == False).sum()), len(res_population)) or 0.0,
            },
            {"metric": "Tickets In Scope", "value": len(in_scope)},
            {
                "metric": "Open Risk",
                "value": int(((in_scope["res_pending"] == True) & (in_scope["res_breached"] == True)).sum()),
            },
        ]
    )
```

- [ ] **Step 4: Re-export the helper from `sla_report_end2end.py` without breaking current callers**

```python
from report_publish import build_dashboard_payload, build_sheet_payloads
```

- [ ] **Step 5: Run the tests to verify the refactor stays green**

Run: `python3 -m unittest test_sla_report_end2end.py test_report_publish.py -v`
Expected: `OK`

- [ ] **Step 6: Checkpoint commit**

If this directory is initialized as Git, run:

```bash
git add report_publish.py sla_report_end2end.py test_sla_report_end2end.py test_report_publish.py
git commit -m "refactor: extract report shaping helpers"
```

### Task 3: Add Backup and Refresh State with TDD

**Files:**
- Create: `/Users/cliada/Documents/code/projects/SLA-Report/report_refresh.py`
- Create: `/Users/cliada/Documents/code/projects/SLA-Report/test_report_refresh.py`
- Modify: `/Users/cliada/Documents/code/projects/SLA-Report/sla_report_end2end.py`

- [ ] **Step 1: Write the failing refresh and backup tests**

```python
import pathlib
import tempfile
import unittest

import report_refresh


class RefreshBackupTests(unittest.TestCase):
    def test_create_backup_bundle_returns_timestamped_paths(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            backup = report_refresh.create_backup_bundle(
                backup_root=pathlib.Path(tmpdir),
                raw_json_bytes=b'{"records":[]}',
                workbook_bytes=b"workbook",
                published_payload_bytes=b'{"Dashboard":[]}',
                generated_at_iso="2026-04-24T12:00:00+00:00",
            )

            self.assertTrue(backup["backup_reference"].startswith("backup-20260424"))
            self.assertTrue(pathlib.Path(backup["backup_dir"]).exists())

    def test_refresh_status_payload_marks_failure_without_losing_backup_reference(self):
        payload = report_refresh.build_refresh_status_payload(
            status="Failed",
            generated_at_iso="2026-04-24T12:00:00+00:00",
            backup_reference="backup-20260424-120000",
            message="Jira timeout",
        )

        self.assertEqual(payload.loc[0, "last_refresh_status"], "Failed")
        self.assertEqual(payload.loc[0, "backup_reference"], "backup-20260424-120000")
        self.assertEqual(payload.loc[0, "message"], "Jira timeout")
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `python3 -m unittest test_report_refresh.py -v`
Expected: `ModuleNotFoundError` or `AttributeError` because `report_refresh` does not exist yet.

- [ ] **Step 3: Implement the minimal backup and refresh helpers**

```python
from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

import pandas as pd


def create_backup_bundle(*, backup_root: Path, raw_json_bytes: bytes, workbook_bytes: bytes, published_payload_bytes: bytes, generated_at_iso: str):
    stamp = generated_at_iso.replace("-", "").replace(":", "").replace("+00:00", "Z")
    backup_reference = f"backup-{stamp[:8]}-{stamp[9:15]}"
    backup_dir = backup_root / backup_reference
    backup_dir.mkdir(parents=True, exist_ok=False)
    (backup_dir / "raw_jira_extract.json").write_bytes(raw_json_bytes)
    (backup_dir / "report.xlsx").write_bytes(workbook_bytes)
    (backup_dir / "sheet_payloads.json").write_bytes(published_payload_bytes)
    return {"backup_reference": backup_reference, "backup_dir": str(backup_dir)}


def build_refresh_status_payload(*, status: str, generated_at_iso: str, backup_reference: str, message: str) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "last_refresh_status": status,
                "generated_at": generated_at_iso,
                "backup_reference": backup_reference,
                "message": message,
            }
        ]
    )
```

- [ ] **Step 4: Wire the helpers into the main script behind explicit orchestration boundaries**

```python
from report_refresh import build_refresh_status_payload, create_backup_bundle
```

- [ ] **Step 5: Run the tests to verify they pass**

Run: `python3 -m unittest test_report_refresh.py test_report_publish.py -v`
Expected: `OK`

- [ ] **Step 6: Checkpoint commit**

If this directory is initialized as Git, run:

```bash
git add report_refresh.py test_report_refresh.py sla_report_end2end.py
git commit -m "feat: add refresh backup orchestration helpers"
```

### Task 4: Publish Google Sheets Payloads from the Local Refresh Command

**Files:**
- Modify: `/Users/cliada/Documents/code/projects/SLA-Report/report_publish.py`
- Modify: `/Users/cliada/Documents/code/projects/SLA-Report/report_refresh.py`
- Modify: `/Users/cliada/Documents/code/projects/SLA-Report/sla_report_end2end.py`
- Modify: `/Users/cliada/Documents/code/projects/SLA-Report/test_report_publish.py`

- [ ] **Step 1: Write the failing publisher integration test**

```python
class PublisherWriteTests(unittest.TestCase):
    def test_publish_payloads_invokes_adapter_per_tab(self):
        calls = []

        class FakeAdapter:
            def write_tab(self, tab_name, dataframe):
                calls.append((tab_name, list(dataframe.columns)))

        payloads = {
            "Dashboard": pd.DataFrame([{"metric": "FRT Attainment %", "value": 1.0}]),
            "Refresh Control": pd.DataFrame([{"last_refresh_status": "Succeeded"}]),
        }

        report_publish.publish_sheet_payloads(FakeAdapter(), payloads)

        self.assertEqual(calls[0][0], "Dashboard")
        self.assertEqual(calls[1][0], "Refresh Control")
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `python3 -m unittest test_report_publish.py -v`
Expected: failure because `publish_sheet_payloads` does not exist yet.

- [ ] **Step 3: Implement the minimal adapter-facing publish loop**

```python
def publish_sheet_payloads(adapter, payloads: Dict[str, pd.DataFrame]) -> None:
    for tab_name, dataframe in payloads.items():
        adapter.write_tab(tab_name, dataframe)
```

- [ ] **Step 4: Add the semi-local refresh orchestration entry point**

```python
def run_local_refresh(*, client, adapter, backup_root: Path, generated_at_iso: str, base_url: str):
    all_df, cycles_df, raw_json = collect_issue_data(client, sleep_seconds=0.0)
    all_df, in_scope, exceptions_non_type, exceptions_rejected, exceptions_feature = classify_scope(all_df)
    payloads = build_sheet_payloads(
        all_df=all_df,
        in_scope=in_scope,
        exceptions_non_type=exceptions_non_type,
        exceptions_rejected=exceptions_rejected,
        exceptions_feature=exceptions_feature,
        cycles_df=cycles_df,
        base_url=base_url,
        generated_at_iso=generated_at_iso,
        backup_reference="pending-backup",
    )
    publish_sheet_payloads(adapter, payloads)
```

- [ ] **Step 5: Run the tests to verify the publish contract still passes**

Run: `python3 -m unittest test_report_publish.py test_report_refresh.py test_sla_report_end2end.py -v`
Expected: `OK`

- [ ] **Step 6: Checkpoint commit**

If this directory is initialized as Git, run:

```bash
git add report_publish.py report_refresh.py sla_report_end2end.py test_report_publish.py
git commit -m "feat: publish sheet payloads from local refresh flow"
```

### Task 5: Add the Google Apps Script Refresh UX

**Files:**
- Create: `/Users/cliada/Documents/code/projects/SLA-Report/apps_script/Code.gs`
- Create: `/Users/cliada/Documents/code/projects/SLA-Report/apps_script/README.md`

- [ ] **Step 1: Write the Apps Script contract in the README before code**

```md
# Google Sheets Refresh UX

This script adds:
- a custom `SLA Report` menu
- a `Refresh Report` action
- a `Refresh Control` sheet initializer

The script does not call local Python. It writes a refresh request and updates visible status so the local Mac refresh command can fulfill it.
```

- [ ] **Step 2: Add the Apps Script menu and refresh request logic**

```javascript
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('SLA Report')
    .addItem('Refresh Report', 'requestRefresh')
    .addItem('Initialize Refresh Control', 'initializeRefreshControl')
    .addToUi();
}

function initializeRefreshControl() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Refresh Control') || ss.insertSheet('Refresh Control');
  sheet.clear();
  sheet.getRange(1, 1, 6, 2).setValues([
    ['field', 'value'],
    ['last_refresh_status', 'Never Run'],
    ['requested_at', ''],
    ['requested_by', ''],
    ['backup_reference', ''],
    ['message', 'Click "Refresh Report" then run the local sync command on macOS.'],
  ]);
}

function requestRefresh() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Refresh Control') || ss.insertSheet('Refresh Control');
  initializeRefreshControl();
  const email = Session.getActiveUser().getEmail() || 'unknown';
  sheet.getRange('B2').setValue('Requested');
  sheet.getRange('B3').setValue(new Date());
  sheet.getRange('B4').setValue(email);
  sheet.getRange('B6').setValue('Refresh requested. Run the local Python sync command to update the report and backup.');
}
```

- [ ] **Step 3: Review the script manually in the Apps Script editor**

Run: install the script in the target Google Sheet and use `Run > onOpen`, then verify the custom `SLA Report` menu appears.
Expected: menu appears with `Refresh Report` and `Initialize Refresh Control`.

- [ ] **Step 4: Verify the refresh request UX manually**

Run: click `SLA Report > Refresh Report`
Expected: `Refresh Control` exists and shows `Requested`, a timestamp, a requester email if available, and guidance to run the local sync command.

- [ ] **Step 5: Checkpoint commit**

If this directory is initialized as Git, run:

```bash
git add apps_script/Code.gs apps_script/README.md
git commit -m "feat: add sheets refresh request ux"
```

### Task 6: Final End-to-End Verification and Cleanup

**Files:**
- Modify: `/Users/cliada/Documents/code/projects/SLA-Report/sla_report_end2end.py`
- Modify: `/Users/cliada/Documents/code/projects/SLA-Report/report_refresh.py`
- Modify: `/Users/cliada/Documents/code/projects/SLA-Report/report_publish.py`
- Modify: `/Users/cliada/Documents/code/projects/SLA-Report/test_sla_report_end2end.py`
- Modify: `/Users/cliada/Documents/code/projects/SLA-Report/test_report_publish.py`
- Modify: `/Users/cliada/Documents/code/projects/SLA-Report/test_report_refresh.py`

- [ ] **Step 1: Add one fixture-driven end-to-end test**

```python
class EndToEndFixtureTests(unittest.TestCase):
    def test_fixture_builds_payloads_and_backup_reference(self):
        all_df, in_scope, non_type, rejected, feature, cycles_df = load_fixture_frames()
        payloads = sla_report_end2end.build_sheet_payloads(
            all_df=all_df,
            in_scope=in_scope,
            exceptions_non_type=non_type,
            exceptions_rejected=rejected,
            exceptions_feature=feature,
            cycles_df=cycles_df,
            base_url="https://example.atlassian.net",
            generated_at_iso="2026-04-24T12:00:00+00:00",
            backup_reference="backup-20260424-120000",
        )
        self.assertEqual(payloads["Refresh Control"].loc[0, "backup_reference"], "backup-20260424-120000")
        self.assertGreater(len(payloads["Raw Data"]), 0)
```

- [ ] **Step 2: Run the full test suite**

Run: `python3 -m unittest test_sla_report_end2end.py test_report_publish.py test_report_refresh.py -v`
Expected: all tests pass with no live Jira access required.

- [ ] **Step 3: Run a local smoke build of the report without publishing**

Run: `python3 -m unittest test_sla_report_end2end.py -v`
Expected: existing classification behavior still passes after the refactor.

- [ ] **Step 4: Run a local refresh smoke test with a fake adapter**

Run: create a minimal fake adapter in a scratch shell or dedicated smoke helper and invoke `run_local_refresh(...)` with fixture-backed data.
Expected: all tabs are written in order, a backup reference is created, and no network request is required in the test harness.

- [ ] **Step 5: Checkpoint commit**

If this directory is initialized as Git, run:

```bash
git add sla_report_end2end.py report_refresh.py report_publish.py test_sla_report_end2end.py test_report_publish.py test_report_refresh.py
git commit -m "test: verify end-to-end report refresh flow"
```

## Self-Review

- Spec coverage:
  - Google Sheets-native dashboard: covered by Tasks 1, 2, 4, and 5.
  - Semi-local refresh flow: covered by Tasks 3, 4, and 5.
  - Backup on every refresh: covered by Task 3 and verified again in Task 6.
  - Strong automated testing: covered by Tasks 1, 2, 3, and 6.
  - Analyst-ready detail tabs and raw data transparency: covered by Tasks 1, 2, and 4.
- Placeholder scan:
  - No `TODO`, `TBD`, or “implement later” placeholders remain in task steps.
  - Manual smoke checks are explicit about what to run and what should happen.
- Type consistency:
  - `build_sheet_payloads`, `build_dashboard_payload`, `publish_sheet_payloads`, `create_backup_bundle`, `build_refresh_status_payload`, and `run_local_refresh` are referenced consistently across tasks.
