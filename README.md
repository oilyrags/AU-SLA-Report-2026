# SLA Report

Python tooling for building the SMG Automotive SLA report from Jira/JSM data, exporting an Excel workbook, and optionally publishing report tabs to Google Sheets.

## Setup

Create the local virtualenv and install dependencies:

```bash
python -m venv .venv
.venv/bin/pip install -r requirements.txt
```

The project scripts intentionally use `.venv/bin/python`. Running tests with the system or Anaconda `python` may fail if dependencies such as `pandas` are not installed there.

## Test

```bash
./scripts/test.sh
```

You can pass normal `unittest` selectors:

```bash
./scripts/test.sh test_report_refresh
```

## Refresh

Run the local refresh from the project root:

```bash
./scripts/refresh_report_local.sh
```

The refresh command:

- checks Jira connectivity
- pulls issues for the selected project/year
- enriches issues with SLA cycles and changelog history
- writes raw JSON and the Excel workbook through temporary sibling files before replacing final artifacts
- creates a backup bundle under `backups/`
- publishes tabs to Google Sheets when Google config is present

## Configuration

Required Jira configuration:

- `JIRA_BASE_URL`
- `JIRA_EMAIL`
- `JIRA_API_TOKEN`

Report scope:

- `JIRA_PROJECT_KEY`, default `ASD`
- `SLA_REPORT_YEAR`, default `2026`
- or CLI flags `--project-key` and `--year`

Google Sheets target:

- `GOOGLE_SHEETS_SPREADSHEET_ID`

Preferred user OAuth:

- `GOOGLE_OAUTH_CLIENT_SECRET_JSON_PATH` or `GOOGLE_OAUTH_CLIENT_SECRET_JSON`
- optional `GOOGLE_OAUTH_TOKEN_PATH`

Service-account fallback:

- `GOOGLE_SERVICE_ACCOUNT_JSON_PATH` or `GOOGLE_SERVICE_ACCOUNT_JSON`

OAuth token cache files and service-account signing temp files are written with owner-only permissions.

## Generated Artifacts

Generated files are intentionally ignored by git:

- `raw_jira_extract_<year>ytd.json`
- `SMG_Automotive_SLA_Report_<year>YTD.xlsx`
- `backups/`

The `archive/` folder is also ignored and intended for local historical artifacts.

## Module Map

- `sla_rules.py`: shared report scope, target tables, issue-type rules, manual overrides, and pure helpers.
- `jira_client.py`: Jira config, prompting, API client, and retry behavior.
- `workbook_writer.py`: Jira issue enrichment, scope classification, narrative generation, and Excel workbook creation.
- `report_publish.py`: Google Sheets payload construction and tab publishing.
- `report_refresh.py`: end-to-end local refresh orchestration, backups, atomic writes, and optional Google publish.
- `sla_report_end2end.py`: compatibility CLI facade for the older single-script workflow.
