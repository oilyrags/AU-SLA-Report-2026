# SLA Report Redesign Design

Date: 2026-04-24
Project: SLA Report
Scope: Redesign the current Python-generated SLA report into a Google Sheets-native command center with a modern executive UI, analyst-friendly detail tabs, semi-local refresh workflow, automated backup, and strong test coverage.

## Goals

- Deliver a state-of-the-art report UI suitable for executive presentation and analyst exploration.
- Preserve Python as the source of truth for Jira extraction, SLA calculations, scope classification, and narrative generation.
- Make Google Sheets the primary presentation surface.
- Add a visible `Refresh Report` control in Google Sheets.
- Support the agreed semi-local refresh model on macOS/GDrive rather than pretending Google Sheets can execute local Python directly.
- Create a backup on each refresh so previous report states can be recovered.
- Build the changes with TDD discipline and meaningful automated verification.

## Non-Goals

- Replacing Jira as the source system.
- Moving SLA calculation logic into Apps Script.
- Building a fully autonomous cloud refresh service in this iteration.
- Maintaining Excel as the primary end-user experience.

## Chosen Product Shape

Google Sheets becomes the main report UX.

Python remains the canonical backend engine and publisher.

The report will include:

- `Dashboard`: premium executive landing page with KPI cards, trend views, narrative blocks, and risk summaries.
- `Summary`: quarterly and YTD SLA attainment tables.
- `Teams`: per-team performance breakdowns.
- `Breaches`: actionable breach details.
- `Trends`: weekly and pattern analysis.
- `Exceptions`: exclusions and rejected-ticket analysis.
- `Methodology`: assumptions, extraction logic, and caveats.
- `Raw Data`: structured published tables used for transparency and debugging.
- `Refresh Control`: status, request marker, last success/failure details, and backup references.

## Visual Direction

The dashboard should feel intentionally designed rather than like a default spreadsheet:

- Strong information hierarchy with a branded hero area.
- KPI cards for FRT attainment, resolution attainment, in-scope volume, and open-risk volume.
- Executive narrative and recommendation blocks near the top.
- Focused visuals for quarter trends, priority health, and team risk.
- Analyst tabs remain dense enough for filtering and review, but receive consistent headers, spacing, and formatting.
- Styling should be optimized for Google Sheets conventions rather than trying to mimic openpyxl chart layouts.

## Architecture

### Python responsibilities

Python owns:

- Jira authentication and extraction.
- Raw SLA cycle collection.
- Scope classification.
- Metric and KPI derivation.
- Narrative generation.
- Transformation of report content into publishable sheet payloads.
- Refresh execution orchestration.
- Backup creation and retention metadata.

The existing monolithic script should be refactored into smaller units with clear responsibilities so they can be tested independently.

### Google Apps Script responsibilities

Apps Script owns:

- Adding a custom menu and/or button entry point for `Refresh Report`.
- Managing a visible refresh request/status area inside the sheet.
- Recording that a refresh was requested.
- Showing operator guidance for the local refresh step.
- Displaying last success/failure timestamps and messages.
- Maintaining light presentation helpers for tab formatting if needed.

Apps Script must not own business rules such as SLA attainment logic, priority targets, or scope classification.

## Refresh Workflow

The agreed refresh model is semi-local:

1. A user clicks `Refresh Report` in Google Sheets.
2. Apps Script updates `Refresh Control` with:
   - requested timestamp
   - requester identity when available
   - status such as `Requested`
   - guidance telling the operator to run the local sync command
3. A local Python refresh command is run on the Mac.
4. Python detects the outstanding refresh request, pulls Jira data, rebuilds all report outputs, publishes the updated tabs to Google Sheets, creates a backup, and updates refresh status.
5. On success, `Refresh Control` shows a success timestamp and backup reference.
6. On failure, `Refresh Control` shows failure state, timestamp, and short error summary without corrupting the previous published report.

## Backup Strategy

Each refresh must create a recoverable backup before overwriting the published report state.

Recommended backup behavior:

- Save a timestamped backup artifact for every refresh run.
- Keep the raw JSON extract as part of the backup set.
- Preserve the previously published structured datasets so the latest good state can be restored.
- Record backup location or backup identifier in `Refresh Control`.
- Apply a simple retention rule, for example keeping the most recent N backups locally.

Practical implementation options:

- Timestamped local export directory containing raw JSON plus published tab payload snapshots.
- Optional archived `.xlsx` output retained alongside the Sheets publish artifacts for historical portability.

Minimum acceptable outcome:

- A failed refresh must not destroy the previous good published state.
- An operator must be able to identify and restore a previous backup without manual reconstruction.

## Data Publishing Model

The current script already computes useful structures for:

- executive dashboard KPIs
- summary blocks
- team analysis
- breach detail
- weekly trends
- narrative output
- exceptions and methodology

Those outputs should be normalized into explicit publish payloads that can be written to Google Sheets tabs. The publishing boundary should be deterministic and testable, with stable column ordering and formatting expectations.

## Testing Strategy

TDD is required for the Python side of the redesign.

### Unit tests

Add or extend tests for:

- scope classification behavior
- summary block generation
- KPI derivation for dashboard cards
- team analysis shaping
- breach table shaping
- trend output shaping
- backup metadata creation
- refresh status transitions
- publish payload schema stability

### Contract tests

Use fixture-based tests to verify the shaped outputs consumed by Google Sheets:

- tab names
- column order
- required metadata fields
- backup references
- refresh control payload

### Integration-style verification

Add at least one fixture-driven end-to-end report build path that exercises:

- raw input fixture
- classification
- KPI derivation
- payload generation
- backup creation

This should not require live Jira for normal test runs.

### Apps Script testing posture

Keep Apps Script intentionally thin so most correctness is proven in Python. For Apps Script, focus on:

- deterministic request/status cell updates
- menu/button wiring
- minimal failure messaging

Where possible, validate behavior through documented manual smoke steps plus contract assumptions rather than migrating core business logic into script code.

## File and Module Changes

Expected additions or changes:

- refactor [`sla_report_end2end.py`](/Users/cliada/Documents/code/projects/SLA-Report/sla_report_end2end.py) into clearer functions and possibly helper modules
- expand [`test_sla_report_end2end.py`](/Users/cliada/Documents/code/projects/SLA-Report/test_sla_report_end2end.py) with more granular coverage
- add Google Sheets publishing support
- add backup creation support
- add refresh control/status support
- add Apps Script project files and usage instructions

## Operational UX

The finished report should support two modes at once:

- Executive mode: open the dashboard and understand performance, risks, and actions in under a minute.
- Analyst mode: move into detail tabs for filters, evidence, outliers, and methodology without losing traceability.

The refresh experience should be operationally honest:

- the sheet should show when data is stale
- the sheet should show whether a refresh is pending, successful, or failed
- the sheet should show where the corresponding backup lives

## Risks and Mitigations

- Google Sheets layout limitations versus bespoke BI tools
  - mitigate by using strong structure, restrained visuals, and clear hierarchy instead of over-designing
- Semi-local refresh depends on a local operator step
  - mitigate by making guidance explicit and status visible in the sheet
- Current script is large and contains duplicated/report-specific logic
  - mitigate by refactoring around tested data-shaping boundaries
- Publish failures could partially update tabs
  - mitigate by creating backups first and writing status explicitly

## Acceptance Criteria

- The primary deliverable is a Google Sheets-native report UI.
- The dashboard looks materially more modern and polished than the current workbook output.
- The report remains useful for both executives and analysts.
- A `Refresh Report` control exists in Google Sheets.
- The refresh workflow updates visible status and requires the agreed local Python step.
- Every refresh creates a usable backup.
- Automated Python tests cover the new shaping and refresh behaviors.
- Normal test execution does not require live Jira credentials.

## Open Assumptions Locked For This Iteration

- Refresh is semi-local, not cloud-run.
- Google Sheets is the primary end-user surface.
- Python remains the single source of truth for business logic.
- Backup is required on every refresh.
- Full Excel feature parity is not required if the Sheets UX is better overall.
