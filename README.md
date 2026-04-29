# SLA Report

Python tooling for building the SMG Automotive SLA report from Jira/JSM data. The refresh workflow pulls Jira issues and SLA cycle data, classifies in-scope and exception tickets, writes an Excel workbook, creates a local backup bundle, and optionally publishes report tabs to Google Sheets.

## What This Project Does

- Pulls Jira/JSM issues for a selected project and report year.
- Enriches tickets with true SLA cycle and changelog history.
- Applies the SLA report scope and exception rules in `sla_rules.py`.
- Generates `SMG_Automotive_SLA_Report_<year>YTD.xlsx`.
- Saves the raw Jira extract as `raw_jira_extract_<year>ytd.json`.
- Creates timestamped backup bundles under `backups/`.
- Optionally publishes tabs to a Google Sheet.
- Provides a lightweight Apps Script menu for requesting refreshes from the sheet.

## Requirements

- Python 3.10. The repository includes `.python-version` with `3.10`.
- Internet access to Jira/JSM.
- Jira credentials with permission to read the target project and SLA data.
- Optional: access to the target Google Sheet and a Google OAuth client JSON or service-account JSON for publishing.

The shell scripts intentionally use `.venv/bin/python`, so create the virtual environment in the project root before running tests or refreshes.

## First-Time Setup

Clone the repository and enter the project directory:

```bash
git clone https://github.com/oilyrags/AU-SLA-Report-2026.git
cd AU-SLA-Report-2026
```

Create the virtual environment and install dependencies:

```bash
python -m venv .venv
.venv/bin/pip install -r requirements.txt
```

Run the test suite:

```bash
./scripts/test.sh
```

You can pass normal `unittest` selectors when you only want part of the suite:

```bash
./scripts/test.sh test_report_refresh
```

## Jira Configuration

The refresh command needs these Jira values. You can export them before running the script, or leave them unset and the script will prompt for them interactively.

```bash
export JIRA_BASE_URL="https://your-domain.atlassian.net"
export JIRA_EMAIL="you@example.com"
export JIRA_API_TOKEN="your-atlassian-api-token"
```

Report scope defaults:

- `JIRA_PROJECT_KEY`: defaults to `ASD`
- `SLA_REPORT_YEAR`: defaults to `2026`

You can also pass scope values as command-line flags:

```bash
./scripts/refresh_report_local.sh --project-key ASD --year 2026
```

## Run A Local Refresh

From the project root:

```bash
./scripts/refresh_report_local.sh
```

The script runs `report_refresh.py`, checks Jira connectivity, pulls issues/SLA cycles/changelogs, writes the workbook and raw JSON extract, creates a backup bundle, and publishes to Google Sheets when Google configuration is present.

Useful options:

```bash
./scripts/refresh_report_local.sh --project-key ASD --year 2026
./scripts/refresh_report_local.sh --sleep-seconds 0.25
```

`--sleep-seconds` controls the delay between per-issue SLA/changelog API calls. Increase it if Jira rate limiting becomes noisy.

## Non-Interactive Refresh Example

```bash
JIRA_BASE_URL="https://your-domain.atlassian.net" \
JIRA_EMAIL="you@example.com" \
JIRA_API_TOKEN="your-atlassian-api-token" \
JIRA_PROJECT_KEY="ASD" \
SLA_REPORT_YEAR="2026" \
./scripts/refresh_report_local.sh
```

With Google Sheets publishing enabled:

```bash
JIRA_BASE_URL="https://your-domain.atlassian.net" \
JIRA_EMAIL="you@example.com" \
JIRA_API_TOKEN="your-atlassian-api-token" \
GOOGLE_SHEETS_SPREADSHEET_ID="your-spreadsheet-id" \
GOOGLE_OAUTH_CLIENT_SECRET_JSON_PATH="/path/to/oauth-client.json" \
./scripts/refresh_report_local.sh --project-key ASD --year 2026
```

## Output Files

A successful local refresh writes:

- `SMG_Automotive_SLA_Report_<year>YTD.xlsx`: the Excel workbook.
- `raw_jira_extract_<year>ytd.json`: the raw Jira/JSM extract used to build the report.
- `backups/backup-<timestamp>/raw_jira_extract.json`: backup copy of the raw extract.
- `backups/backup-<timestamp>/report.xlsx`: backup copy of the workbook.
- `backups/backup-<timestamp>/sheet_payloads.json`: backup copy of the Google Sheets payloads.

These generated files are intentionally ignored by git. The `archive/` folder is also ignored and is intended for local historical artifacts only.

## Optional Google Sheets Publishing

Set `GOOGLE_SHEETS_SPREADSHEET_ID` to the target spreadsheet ID. Then choose one authentication method.

Preferred user OAuth flow:

```bash
export GOOGLE_SHEETS_SPREADSHEET_ID="your-spreadsheet-id"
export GOOGLE_OAUTH_CLIENT_SECRET_JSON_PATH="/path/to/oauth-client.json"
```

You can provide the OAuth client JSON inline instead:

```bash
export GOOGLE_OAUTH_CLIENT_SECRET_JSON='{"installed":{...}}'
```

Optional token cache location:

```bash
export GOOGLE_OAUTH_TOKEN_PATH="$HOME/.config/sla-report/google-oauth-token.json"
```

Service-account fallback:

```bash
export GOOGLE_SHEETS_SPREADSHEET_ID="your-spreadsheet-id"
export GOOGLE_SERVICE_ACCOUNT_JSON_PATH="/path/to/service-account.json"
```

Or provide the service-account JSON inline:

```bash
export GOOGLE_SERVICE_ACCOUNT_JSON='{"type":"service_account",...}'
```

If no Google Sheets configuration is present, the refresh still creates the local workbook, raw JSON, and backup bundle, then logs that Google publishing was skipped.

OAuth token cache files and temporary service-account signing files are written with owner-only permissions.

## Google Sheet Menu

The `apps_script/` folder contains an optional Apps Script layer for the target Google Sheet. It adds a `SLA Report` menu with:

- `Initialize Refresh Control`
- `Refresh Report`

The Apps Script does not call Jira or run Python. It records a refresh request in a `Refresh Control` tab so the local operator knows a refresh was requested.

Install and usage details are in `apps_script/README.md`.

## Published Google Sheet Tabs

When publishing is enabled, the Python refresh writes these tabs:

- `Dashboard`
- `Summary`
- `Teams`
- `Breaches`
- `Trends`
- `Exceptions`
- `Methodology`
- `Raw Data`
- `Refresh Control`

## Troubleshooting

If `./scripts/test.sh` or `./scripts/refresh_report_local.sh` reports a missing virtualenv, recreate it:

```bash
python -m venv .venv
.venv/bin/pip install -r requirements.txt
```

If Jira authentication fails, verify `JIRA_BASE_URL`, `JIRA_EMAIL`, and `JIRA_API_TOKEN`. Atlassian API tokens are not the same thing as your normal account password.

If the Google OAuth browser flow does not complete, confirm that the OAuth client is a Desktop app client and that the signed-in Google account can access the spreadsheet.

If a cached Google token becomes invalid, set `GOOGLE_OAUTH_TOKEN_PATH` to a new path or remove the old token cache file and rerun the refresh.

If Jira rate limiting appears during refreshes, rerun with a larger delay:

```bash
./scripts/refresh_report_local.sh --sleep-seconds 0.5
```

## Development Notes

Run all tests before publishing changes:

```bash
./scripts/test.sh
```

Primary modules:

- `sla_rules.py`: report scope, targets, issue-type rules, manual overrides, and pure helpers.
- `jira_client.py`: Jira config, prompting, API client, pagination, and retry behavior.
- `workbook_writer.py`: issue enrichment, scope classification, narrative generation, and Excel workbook creation.
- `report_publish.py`: Google Sheets payload construction and tab publishing.
- `report_refresh.py`: end-to-end local refresh orchestration, backups, atomic writes, and optional Google publishing.
- `google_sheets_adapter.py`: Google OAuth/service-account auth and Sheets API reads/writes.
- `sla_report_end2end.py`: compatibility CLI facade for the older single-script workflow.

Do not commit credentials, raw Jira extracts, generated workbooks, backup bundles, local virtualenvs, or local scratch files.
