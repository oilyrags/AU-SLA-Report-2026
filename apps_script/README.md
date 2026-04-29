# Google Sheets Refresh UX

This folder contains the thin Google Apps Script layer for the SLA report.

What it adds:
- a custom `SLA Report` menu in the bound Google Sheet
- a `Refresh Report` action that records a refresh request
- an `Initialize Refresh Control` action that seeds the control sheet
- a visible `Refresh Control` tab with request metadata and operator guidance

What it does not do:
- it does not run SLA calculations
- it does not call Jira
- it does not execute the local Python sync command

## Install

1. Open the target Google Sheet.
2. Go to **Extensions > Apps Script**.
3. Create or paste in `Code.gs` from this folder.
4. Save the project and authorize it when prompted.
5. Reload the spreadsheet so the custom menu appears.

## Use

1. Open the sheet and choose **SLA Report > Initialize Refresh Control** once if the control tab has not been created yet.
2. Choose **SLA Report > Refresh Report** when you want to request a refresh.
3. The script will update the `Refresh Control` sheet with:
   - `last_refresh_status = Requested`
   - `requested_at`
   - `requested_by` when Google exposes the active user email
   - the latest `backup_reference` if one is already present
   - guidance to run `./scripts/refresh_report_local.sh`
4. Run `./scripts/refresh_report_local.sh` on the Mac that owns the refresh workflow.

## Local refresh command

Run this from the project root on the Mac that has Jira access:

```bash
./scripts/refresh_report_local.sh
```

The script:
- uses the repo virtualenv at `.venv/bin/python`
- runs [`report_refresh.py`](/Users/cliada/Documents/code/projects/SLA-Report/report_refresh.py)
- prompts for `JIRA_BASE_URL`, `JIRA_EMAIL`, and `JIRA_API_TOKEN` if they are not already exported
- rebuilds the local workbook and raw JSON artifacts
- preserves `requested_at` and `requested_by` in `Refresh Control` when the Python Google Sheets adapter can read the existing tab
- publishes to Google Sheets when `GOOGLE_SHEETS_SPREADSHEET_ID` and Google auth config are present

## Optional Google Sheets publish

Preferred for Okta-backed Google Workspace login:

- `GOOGLE_OAUTH_CLIENT_SECRET_JSON_PATH` pointing to a Google OAuth desktop client JSON file
- `GOOGLE_OAUTH_CLIENT_SECRET_JSON` containing that JSON inline
- optional `GOOGLE_OAUTH_TOKEN_PATH` to control where the cached user token is stored

Fallback for organizations that allow service accounts:

- `GOOGLE_SERVICE_ACCOUNT_JSON_PATH` pointing to a service-account JSON file
- `GOOGLE_SERVICE_ACCOUNT_JSON` containing the service-account JSON inline

Also set `GOOGLE_SHEETS_SPREADSHEET_ID` to the target spreadsheet id.

For the Okta-compatible OAuth flow:

1. Create a Google OAuth client in Google Cloud for a Desktop app.
2. Download the client JSON.
3. Run the refresh command with `GOOGLE_OAUTH_CLIENT_SECRET_JSON_PATH`.
4. On the first run, a browser window opens to Google sign-in.
5. Sign in with your normal Google Workspace account. If your org uses Okta SSO, Google will hand off to Okta during this login.
6. After consent, the local command stores a refresh token and future runs reuse it.

## Test

Run the Python suite from the project root with:

```bash
./scripts/test.sh
```

You can also run it non-interactively:

```bash
JIRA_BASE_URL="https://your-domain.atlassian.net" \
JIRA_EMAIL="you@example.com" \
JIRA_API_TOKEN="your-token" \
GOOGLE_SHEETS_SPREADSHEET_ID="your-spreadsheet-id" \
GOOGLE_OAUTH_CLIENT_SECRET_JSON_PATH="/path/to/oauth-client.json" \
./scripts/refresh_report_local.sh
```

## Control sheet layout

The script keeps the `Refresh Control` sheet in a simple two-column layout:

- `field`
- `value`

Rows include:
- `last_refresh_status`
- `requested_at`
- `requested_by`
- `backup_reference`
- `message`

## Notes

- `requested_by` may show `Not available` if Google Sheets does not expose a user email in the current context.
- The control sheet is intentionally lightweight so the Python side remains the source of truth for report data and refresh behavior.
