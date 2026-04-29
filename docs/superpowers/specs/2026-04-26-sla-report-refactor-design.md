# SLA Report Refactor Design

## Goal

Implement the full recommendation set from the code analysis while preserving the existing report behavior and current command-line workflows.

The work should reduce duplicated business rules, make the refresh flow safer, improve credential handling, add a clearer developer entry point, and split the largest responsibilities into smaller modules without rewriting the workbook design from scratch.

## Scope

- Extract shared SLA/report configuration into `sla_rules.py`.
- Extract Jira API/client and credential prompting into `jira_client.py`.
- Extract workbook and narrative generation into `workbook_writer.py`.
- Keep `sla_report_end2end.py` as a compatibility CLI facade and import surface for existing tests/users.
- Update publishing and refresh code to use shared rules/config.
- Add report scope parameters for project key and year via CLI and environment variables.
- Harden Google OAuth token and temporary private-key file permissions.
- Preserve Google Sheets refresh-request audit fields when publishing refresh status.
- Make local report artifact writes atomic where practical.
- Add root documentation and a repeatable test script.
- Add tests for new shared rules/config, metadata preservation, and credential-file hardening.

## Architecture

`sla_rules.py` owns report constants and small pure helpers:

- default project key and year
- issue type names and aliases
- manual include/exclude key sets
- feature/bug cue sets
- priorities and target table
- date helpers, target lookup, percent calculation, percentile calculation, and value extraction helpers

`jira_client.py` owns Jira-specific behavior:

- `JiraConfig`
- `JiraClient`
- retry behavior
- `prompt_env`

`workbook_writer.py` owns report transformation output:

- `classify_scope`
- `collect_issue_data`
- `narrative_from_data`
- workbook styling and `build_workbook`

`sla_report_end2end.py` remains a thin entry point that re-exports the compatibility names used by tests and callers, then delegates CLI execution to the extracted modules.

`report_publish.py` imports priorities and target lookup from `sla_rules.py` so published sheets and Excel sheets cannot drift.

## Data Flow

The refresh flow remains:

1. Read Jira config from CLI/env prompts.
2. Query Jira issues for the selected project/year.
3. Enrich each issue with JSM SLA cycles and Jira changelog.
4. Classify scope using shared rules.
5. Write raw JSON and workbook artifacts.
6. Build Google Sheets payloads.
7. Create a backup bundle.
8. Publish payloads to Google Sheets when configured.

The user-facing commands remain:

- `./scripts/refresh_report_local.sh`
- `.venv/bin/python report_refresh.py`
- `.venv/bin/python sla_report_end2end.py`
- `./scripts/test.sh`

## Error Handling

- Jira and Sheets API errors continue raising explicit `RuntimeError`s with HTTP context.
- Raw JSON and workbook output should be written through temporary sibling files and atomically replaced.
- Backup creation remains fail-fast so a refresh cannot claim a backup reference that was not written.
- If Sheets publishing fails, Python should publish a failure refresh-control payload when possible.
- Failure status should retain the same backup reference and preserve request metadata if known.

## Security

- OAuth token cache writes should create parent directories and token files with owner-only permissions.
- Service-account private keys should be written to temp files with owner-only permissions and removed after signing.
- Existing env-var auth options remain supported.

## Testing

Keep the current `unittest` suite passing and add focused coverage for:

- shared target lookup is used by publish helpers
- CLI/env report scope parsing for project key/year
- refresh-control request metadata preservation
- token cache file mode
- private-key temp file mode where testable without invoking real Google APIs

Verification command:

```bash
./scripts/test.sh
```

Fallback:

```bash
.venv/bin/python -m unittest
```

## Non-Goals

- No visual redesign of the Excel workbook.
- No migration to a package build system beyond lightweight developer docs/scripts.
- No live Jira or Google API calls in tests.
- No destructive cleanup of existing generated artifacts.
