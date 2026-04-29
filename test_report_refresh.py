#!/usr/bin/env python3
import contextlib
import io
import pathlib
import tempfile
import unittest

import pandas as pd
import report_refresh


class RefreshBackupTests(unittest.TestCase):
    def test_run_local_refresh_builds_payloads_and_publishes_them(self):
        calls = []

        original_collect_issue_data = report_refresh.collect_issue_data
        original_classify_scope = report_refresh.classify_scope
        original_build_refresh_payloads = report_refresh.build_refresh_payloads_from_frames
        original_publish_sheet_payloads = report_refresh.publish_sheet_payloads

        try:
            def fake_collect_issue_data(client, sleep_seconds):
                calls.append(("collect_issue_data", client, sleep_seconds))
                return "all_df", "cycles_df", {"raw": True}

            def fake_classify_scope(all_df):
                calls.append(("classify_scope", all_df))
                return "all_df", "in_scope", "exceptions_non_type", "exceptions_rejected", "exceptions_feature"

            def fake_build_sheet_payloads(**kwargs):
                calls.append(("build_sheet_payloads", kwargs))
                return {"Dashboard": "dashboard-frame", "Refresh Control": "refresh-frame"}

            def fake_publish_sheet_payloads(adapter, payloads):
                calls.append(("publish_sheet_payloads", adapter, payloads))

            report_refresh.collect_issue_data = fake_collect_issue_data
            report_refresh.classify_scope = fake_classify_scope
            report_refresh.build_refresh_payloads_from_frames = fake_build_sheet_payloads
            report_refresh.publish_sheet_payloads = fake_publish_sheet_payloads

            client = object()
            adapter = object()
            result = report_refresh.run_local_refresh(
                client=client,
                adapter=adapter,
                generated_at_iso="2026-04-24T12:00:00+00:00",
                base_url="https://example.atlassian.net",
            )

            self.assertEqual(result, {"Dashboard": "dashboard-frame", "Refresh Control": "refresh-frame"})
            self.assertEqual([call[0] for call in calls], [
                "collect_issue_data",
                "classify_scope",
                "build_sheet_payloads",
                "publish_sheet_payloads",
            ])
            self.assertEqual(calls[0][2], 0.0)
            self.assertEqual(calls[2][1]["generated_at_iso"], "2026-04-24T12:00:00+00:00")
            self.assertEqual(calls[2][1]["base_url"], "https://example.atlassian.net")
            self.assertEqual(calls[2][1]["backup_reference"], "pending-backup")
            self.assertIs(calls[3][1], adapter)
        finally:
            report_refresh.collect_issue_data = original_collect_issue_data
            report_refresh.classify_scope = original_classify_scope
            report_refresh.build_refresh_payloads_from_frames = original_build_refresh_payloads
            report_refresh.publish_sheet_payloads = original_publish_sheet_payloads

    def test_refresh_report_local_skips_google_publish_when_config_is_absent(self):
        calls = []

        original_collect_issue_data = report_refresh.collect_issue_data
        original_classify_scope = report_refresh.classify_scope
        original_build_workbook = report_refresh.build_workbook
        original_build_refresh_payloads = report_refresh.build_refresh_payloads_from_frames
        original_publish_sheet_payloads = report_refresh.publish_sheet_payloads
        original_build_google_sheets_adapter_from_env = report_refresh.build_google_sheets_adapter_from_env

        try:
            def fake_collect_issue_data(client, sleep_seconds):
                calls.append(("collect_issue_data", client, sleep_seconds))
                return "all_df", "cycles_df", {"raw": True}

            def fake_classify_scope(all_df):
                calls.append(("classify_scope", all_df))
                return "all_df", "in_scope", "exceptions_non_type", "exceptions_rejected", "exceptions_feature"

            def fake_build_workbook(**kwargs):
                calls.append(("build_workbook", kwargs))
                pathlib.Path(kwargs["output_path"]).write_bytes(b"workbook-bytes")
                return {"pulled": 1, "in_scope": 1, "exceptions_total": 0, "exceptions_non_type": 0, "exceptions_rejected": 0, "exceptions_feature": 0}

            def fake_build_sheet_payloads(**kwargs):
                calls.append(("build_sheet_payloads", kwargs))
                return {"Dashboard": "dashboard-frame"}

            def fake_publish_sheet_payloads(adapter, payloads):
                calls.append(("publish_sheet_payloads", adapter, payloads))

            report_refresh.collect_issue_data = fake_collect_issue_data
            report_refresh.classify_scope = fake_classify_scope
            report_refresh.build_workbook = fake_build_workbook
            report_refresh.build_refresh_payloads_from_frames = fake_build_sheet_payloads
            report_refresh.publish_sheet_payloads = fake_publish_sheet_payloads
            report_refresh.build_google_sheets_adapter_from_env = lambda: None

            with tempfile.TemporaryDirectory() as tmpdir:
                tmpdir_path = pathlib.Path(tmpdir)
                stderr = io.StringIO()
                with contextlib.redirect_stderr(stderr):
                    result = report_refresh.refresh_report_local(
                        client=object(),
                        base_url="https://example.atlassian.net",
                        generated_at_iso="2026-04-24T12:00:00Z",
                        output_path=tmpdir_path / "report.xlsx",
                        raw_json_path=tmpdir_path / "raw.json",
                        sleep_seconds=0.0,
                    )

                self.assertEqual(result["google_published"], False)
                self.assertIn("skipping Google Sheets publish", stderr.getvalue())
                self.assertTrue((tmpdir_path / "raw.json").exists())
                self.assertEqual((tmpdir_path / "raw.json").read_text(encoding="utf-8"), '{\n  "raw": true\n}')
                self.assertEqual([call[0] for call in calls], [
                    "collect_issue_data",
                    "classify_scope",
                    "build_workbook",
                    "build_sheet_payloads",
                ])
                self.assertNotIn("publish_sheet_payloads", [call[0] for call in calls])
                self.assertEqual(result["output_path"], str(tmpdir_path / "report.xlsx"))
                self.assertEqual(result["raw_json_path"], str(tmpdir_path / "raw.json"))
        finally:
            report_refresh.collect_issue_data = original_collect_issue_data
            report_refresh.classify_scope = original_classify_scope
            report_refresh.build_workbook = original_build_workbook
            report_refresh.build_refresh_payloads_from_frames = original_build_refresh_payloads
            report_refresh.publish_sheet_payloads = original_publish_sheet_payloads
            report_refresh.build_google_sheets_adapter_from_env = original_build_google_sheets_adapter_from_env

    def test_create_backup_bundle_returns_timestamped_paths(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            backup = report_refresh.create_backup_bundle(
                backup_root=pathlib.Path(tmpdir),
                raw_json_bytes=b'{"records":[]}',
                workbook_bytes=b"workbook",
                published_payload_bytes=b'{"Dashboard":[]}',
                generated_at_iso="2026-04-24T12:00:00+00:00",
            )

            backup_dir = pathlib.Path(backup["backup_dir"])
            self.assertTrue(backup["backup_reference"].startswith("backup-20260424"))
            self.assertTrue(backup_dir.exists())
            self.assertEqual((backup_dir / "raw_jira_extract.json").read_bytes(), b'{"records":[]}')
            self.assertEqual((backup_dir / "report.xlsx").read_bytes(), b"workbook")
            self.assertEqual((backup_dir / "sheet_payloads.json").read_bytes(), b'{"Dashboard":[]}')

    def test_refresh_status_payload_marks_failure_without_losing_backup_reference(self):
        payload = report_refresh.build_refresh_status_payload(
            status="Failed",
            generated_at_iso="2026-04-24T12:00:00+00:00",
            backup_reference="backup-20260424-120000",
            message="Jira timeout",
            request_metadata={
                "requested_at": "2026-04-24 13:00:00 CEST",
                "requested_by": "operator@example.com",
            },
        )

        fields = payload.set_index("field")["value"].to_dict()
        self.assertEqual(payload.columns.tolist(), ["field", "value"])
        self.assertEqual(fields["last_refresh_status"], "Failed")
        self.assertEqual(fields["generated_at"], "2026-04-24T12:00:00+00:00")
        self.assertEqual(fields["backup_reference"], "backup-20260424-120000")
        self.assertEqual(fields["message"], "Jira timeout")
        self.assertEqual(fields["requested_at"], "2026-04-24 13:00:00 CEST")
        self.assertEqual(fields["requested_by"], "operator@example.com")

    def test_get_refresh_request_metadata_reads_existing_control_values(self):
        class FakeAdapter:
            def read_tab_values(self, tab_name):
                self.tab_name = tab_name
                return [
                    ["field", "value"],
                    ["last_refresh_status", "Requested"],
                    ["requested_at", "2026-04-24 13:00:00 CEST"],
                    ["requested_by", "operator@example.com"],
                    ["message", "Refresh requested"],
                ]

        adapter = FakeAdapter()

        metadata = report_refresh.get_refresh_request_metadata(adapter)

        self.assertEqual(adapter.tab_name, "Refresh Control")
        self.assertEqual(
            metadata,
            {
                "requested_at": "2026-04-24 13:00:00 CEST",
                "requested_by": "operator@example.com",
            },
        )

    def test_refresh_report_local_creates_backup_and_publishes_real_reference(self):
        calls = []

        original_collect_issue_data = report_refresh.collect_issue_data
        original_classify_scope = report_refresh.classify_scope
        original_build_workbook = report_refresh.build_workbook
        original_build_refresh_payloads = report_refresh.build_refresh_payloads_from_frames
        original_publish_sheet_payloads = report_refresh.publish_sheet_payloads

        try:
            def fake_collect_issue_data(client, sleep_seconds):
                return "all_df", "cycles_df", {"raw": True}

            def fake_classify_scope(all_df):
                return "all_df", "in_scope", "exceptions_non_type", "exceptions_rejected", "exceptions_feature"

            def fake_build_workbook(**kwargs):
                pathlib.Path(kwargs["output_path"]).write_bytes(b"workbook-bytes")
                return {"pulled": 1, "in_scope": 1, "exceptions_total": 0, "exceptions_non_type": 0, "exceptions_rejected": 0, "exceptions_feature": 0}

            def fake_build_sheet_payloads(**kwargs):
                calls.append(("build_sheet_payloads", kwargs))
                return {
                    "Refresh Control": pd.DataFrame(
                        [{"field": "backup_reference", "value": kwargs["backup_reference"]}]
                    )
                }

            def fake_publish_sheet_payloads(adapter, payloads):
                calls.append(("publish_sheet_payloads", adapter, payloads))

            report_refresh.collect_issue_data = fake_collect_issue_data
            report_refresh.classify_scope = fake_classify_scope
            report_refresh.build_workbook = fake_build_workbook
            report_refresh.build_refresh_payloads_from_frames = fake_build_sheet_payloads
            report_refresh.publish_sheet_payloads = fake_publish_sheet_payloads

            with tempfile.TemporaryDirectory() as tmpdir:
                tmpdir_path = pathlib.Path(tmpdir)
                result = report_refresh.refresh_report_local(
                    client=object(),
                    base_url="https://example.atlassian.net",
                    generated_at_iso="2026-04-24T12:00:00.123456Z",
                    output_path=tmpdir_path / "report.xlsx",
                    raw_json_path=tmpdir_path / "raw.json",
                    backup_root=tmpdir_path / "backups",
                    sleep_seconds=0.0,
                    google_adapter=object(),
                )

                backup_reference = result["backup_reference"]
                self.assertTrue(backup_reference.startswith("backup-20260424-120000"))
                self.assertNotEqual(backup_reference, "pending-backup")
                self.assertEqual(calls[0][1]["backup_reference"], backup_reference)
                self.assertEqual(calls[1][2]["Refresh Control"].loc[0, "value"], backup_reference)
                self.assertTrue(pathlib.Path(result["backup_dir"]).exists())
                self.assertEqual((pathlib.Path(result["backup_dir"]) / "report.xlsx").read_bytes(), b"workbook-bytes")
                self.assertIn(backup_reference, (pathlib.Path(result["backup_dir"]) / "sheet_payloads.json").read_text(encoding="utf-8"))
        finally:
            report_refresh.collect_issue_data = original_collect_issue_data
            report_refresh.classify_scope = original_classify_scope
            report_refresh.build_workbook = original_build_workbook
            report_refresh.build_refresh_payloads_from_frames = original_build_refresh_payloads
            report_refresh.publish_sheet_payloads = original_publish_sheet_payloads

    def test_refresh_report_local_writes_outputs_through_temp_paths(self):
        workbook_paths = []

        original_collect_issue_data = report_refresh.collect_issue_data
        original_classify_scope = report_refresh.classify_scope
        original_build_workbook = report_refresh.build_workbook
        original_build_refresh_payloads = report_refresh.build_refresh_payloads_from_frames

        try:
            report_refresh.collect_issue_data = lambda client, sleep_seconds: ("all_df", "cycles_df", {"raw": True})
            report_refresh.classify_scope = lambda all_df: ("all_df", "in_scope", "exceptions_non_type", "exceptions_rejected", "exceptions_feature")

            def fake_build_workbook(**kwargs):
                output_path = pathlib.Path(kwargs["output_path"])
                workbook_paths.append(output_path)
                output_path.write_bytes(b"workbook-through-temp")
                return {"pulled": 1, "in_scope": 1, "exceptions_total": 0, "exceptions_non_type": 0, "exceptions_rejected": 0, "exceptions_feature": 0}

            report_refresh.build_workbook = fake_build_workbook
            report_refresh.build_refresh_payloads_from_frames = lambda **kwargs: {"Dashboard": pd.DataFrame([{"metric": "Tickets", "value": 1}])}

            with tempfile.TemporaryDirectory() as tmpdir:
                tmpdir_path = pathlib.Path(tmpdir)
                output_path = tmpdir_path / "report.xlsx"
                raw_json_path = tmpdir_path / "raw.json"

                result = report_refresh.refresh_report_local(
                    client=object(),
                    base_url="https://example.atlassian.net",
                    generated_at_iso="2026-04-24T12:00:00.123456Z",
                    output_path=output_path,
                    raw_json_path=raw_json_path,
                    backup_root=tmpdir_path / "backups",
                    sleep_seconds=0.0,
                    google_adapter=None,
                )

                self.assertEqual(workbook_paths, [tmpdir_path / ".report.xlsx.tmp"])
                self.assertEqual(output_path.read_bytes(), b"workbook-through-temp")
                self.assertEqual(raw_json_path.read_text(encoding="utf-8"), '{\n  "raw": true\n}')
                self.assertFalse((tmpdir_path / ".report.xlsx.tmp").exists())
                self.assertFalse((tmpdir_path / ".raw.json.tmp").exists())
                self.assertEqual(result["output_path"], str(output_path))
        finally:
            report_refresh.collect_issue_data = original_collect_issue_data
            report_refresh.classify_scope = original_classify_scope
            report_refresh.build_workbook = original_build_workbook
            report_refresh.build_refresh_payloads_from_frames = original_build_refresh_payloads

    def test_refresh_report_local_publishes_failure_status_when_sheet_publish_fails(self):
        publish_calls = []

        original_collect_issue_data = report_refresh.collect_issue_data
        original_classify_scope = report_refresh.classify_scope
        original_build_workbook = report_refresh.build_workbook
        original_build_refresh_payloads = report_refresh.build_refresh_payloads_from_frames
        original_publish_sheet_payloads = report_refresh.publish_sheet_payloads

        try:
            report_refresh.collect_issue_data = lambda client, sleep_seconds: ("all_df", "cycles_df", {"raw": True})
            report_refresh.classify_scope = lambda all_df: ("all_df", "in_scope", "exceptions_non_type", "exceptions_rejected", "exceptions_feature")

            def fake_build_workbook(**kwargs):
                pathlib.Path(kwargs["output_path"]).write_bytes(b"workbook-bytes")
                return {"pulled": 1, "in_scope": 1, "exceptions_total": 0, "exceptions_non_type": 0, "exceptions_rejected": 0, "exceptions_feature": 0}

            def fake_build_sheet_payloads(**kwargs):
                return {"Dashboard": pd.DataFrame([{"metric": "Tickets", "value": 1}])}

            def fake_publish_sheet_payloads(adapter, payloads):
                publish_calls.append(payloads)
                if len(publish_calls) == 1:
                    raise RuntimeError("Sheets outage")

            report_refresh.build_workbook = fake_build_workbook
            report_refresh.build_refresh_payloads_from_frames = fake_build_sheet_payloads
            report_refresh.publish_sheet_payloads = fake_publish_sheet_payloads

            with tempfile.TemporaryDirectory() as tmpdir:
                tmpdir_path = pathlib.Path(tmpdir)
                with self.assertRaisesRegex(RuntimeError, "Sheets outage"):
                    report_refresh.refresh_report_local(
                        client=object(),
                        base_url="https://example.atlassian.net",
                        generated_at_iso="2026-04-24T12:00:00.123456Z",
                        output_path=tmpdir_path / "report.xlsx",
                        raw_json_path=tmpdir_path / "raw.json",
                        backup_root=tmpdir_path / "backups",
                        sleep_seconds=0.0,
                        google_adapter=object(),
                    )

            self.assertEqual(len(publish_calls), 2)
            self.assertEqual(list(publish_calls[1].keys()), ["Refresh Control"])
            fields = publish_calls[1]["Refresh Control"].set_index("field")["value"].to_dict()
            self.assertEqual(fields["last_refresh_status"], "Failed")
            self.assertIn("Sheets outage", fields["message"])
            self.assertTrue(str(fields["backup_reference"]).startswith("backup-20260424-120000"))
        finally:
            report_refresh.collect_issue_data = original_collect_issue_data
            report_refresh.classify_scope = original_classify_scope
            report_refresh.build_workbook = original_build_workbook
            report_refresh.build_refresh_payloads_from_frames = original_build_refresh_payloads
            report_refresh.publish_sheet_payloads = original_publish_sheet_payloads

    def test_create_backup_bundle_uses_higher_resolution_timestamp_within_same_second(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            root = pathlib.Path(tmpdir)
            first = report_refresh.create_backup_bundle(
                backup_root=root,
                raw_json_bytes=b"{}",
                workbook_bytes=b"workbook-1",
                published_payload_bytes=b"payload-1",
                generated_at_iso="2026-04-24T12:00:00.123456+00:00",
            )
            second = report_refresh.create_backup_bundle(
                backup_root=root,
                raw_json_bytes=b"{}",
                workbook_bytes=b"workbook-2",
                published_payload_bytes=b"payload-2",
                generated_at_iso="2026-04-24T12:00:00.789012+00:00",
            )

            self.assertNotEqual(first["backup_reference"], second["backup_reference"])
            self.assertNotEqual(first["backup_dir"], second["backup_dir"])
            self.assertTrue(first["backup_reference"].startswith("backup-20260424"))
            self.assertTrue(second["backup_reference"].startswith("backup-20260424"))


if __name__ == "__main__":
    unittest.main()
