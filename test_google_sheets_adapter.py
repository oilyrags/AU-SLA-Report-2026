#!/usr/bin/env python3
import json
import pathlib
import stat
import tempfile
import time
import unittest
from unittest import mock

import pandas as pd

import google_sheets_adapter


class GoogleSheetsAdapterConfigTests(unittest.TestCase):
    def test_load_google_sheets_adapter_from_env_prefers_user_oauth_when_present(self):
        oauth_client = {
            "installed": {
                "client_id": "oauth-client-id.apps.googleusercontent.com",
                "client_secret": "oauth-secret",
                "auth_uri": "https://accounts.google.com/o/oauth2/v2/auth",
                "token_uri": "https://oauth2.googleapis.com/token",
            }
        }
        service_account = {
            "client_email": "svc-account@example.iam.gserviceaccount.com",
            "private_key": "-----BEGIN PRIVATE KEY-----\nFAKE\n-----END PRIVATE KEY-----\n",
            "token_uri": "https://oauth2.googleapis.com/token",
        }

        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir_path = pathlib.Path(tmpdir)
            oauth_path = tmpdir_path / "oauth.json"
            service_path = tmpdir_path / "service-account.json"
            oauth_path.write_text(json.dumps(oauth_client), encoding="utf-8")
            service_path.write_text(json.dumps(service_account), encoding="utf-8")

            adapter = google_sheets_adapter.load_google_sheets_adapter_from_env(
                env={
                    "GOOGLE_SHEETS_SPREADSHEET_ID": "spreadsheet-123",
                    "GOOGLE_OAUTH_CLIENT_SECRET_JSON_PATH": str(oauth_path),
                    "GOOGLE_SERVICE_ACCOUNT_JSON_PATH": str(service_path),
                    "GOOGLE_OAUTH_TOKEN_PATH": str(tmpdir_path / "token.json"),
                }
            )

            self.assertIsNotNone(adapter)
            self.assertIsInstance(adapter.auth_provider, google_sheets_adapter.UserOAuthTokenProvider)
            self.assertEqual(adapter.config.client_id, "oauth-client-id.apps.googleusercontent.com")
            self.assertEqual(adapter.config.spreadsheet_id, "spreadsheet-123")

    def test_load_google_sheets_adapter_from_env_falls_back_to_service_account(self):
        service_account = {
            "client_email": "svc-account@example.iam.gserviceaccount.com",
            "private_key": "-----BEGIN PRIVATE KEY-----\nFAKE\n-----END PRIVATE KEY-----\n",
            "token_uri": "https://oauth2.googleapis.com/token",
        }

        with tempfile.TemporaryDirectory() as tmpdir:
            sa_path = pathlib.Path(tmpdir) / "service-account.json"
            sa_path.write_text(json.dumps(service_account), encoding="utf-8")

            adapter = google_sheets_adapter.load_google_sheets_adapter_from_env(
                env={
                    "GOOGLE_SHEETS_SPREADSHEET_ID": "spreadsheet-123",
                    "GOOGLE_SERVICE_ACCOUNT_JSON_PATH": str(sa_path),
                }
            )

            self.assertIsNotNone(adapter)
            self.assertIsInstance(adapter.auth_provider, google_sheets_adapter.ServiceAccountTokenProvider)
            self.assertEqual(adapter.config.client_email, service_account["client_email"])


class GoogleSheetsAdapterOAuthTests(unittest.TestCase):
    def test_user_oauth_token_provider_uses_cached_access_token(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            token_path = pathlib.Path(tmpdir) / "token.json"
            token_path.write_text(
                json.dumps(
                    {
                        "access_token": "cached-token",
                        "refresh_token": "refresh-token",
                        "expires_at": time.time() + 3600,
                    }
                ),
                encoding="utf-8",
            )

            config = google_sheets_adapter.UserOAuthConfig(
                spreadsheet_id="spreadsheet-123",
                client_id="client-id",
                client_secret=None,
                auth_uri="https://accounts.google.com/o/oauth2/v2/auth",
                token_uri="https://oauth2.googleapis.com/token",
                token_cache_path=token_path,
            )
            provider = google_sheets_adapter.UserOAuthTokenProvider(config)

            self.assertEqual(provider.get_access_token(), "cached-token")

    def test_user_oauth_token_provider_builds_auth_url_with_pkce(self):
        config = google_sheets_adapter.UserOAuthConfig(
            spreadsheet_id="spreadsheet-123",
            client_id="client-id",
            client_secret="secret",
            auth_uri="https://accounts.google.com/o/oauth2/v2/auth",
            token_uri="https://oauth2.googleapis.com/token",
            token_cache_path=pathlib.Path("/tmp/google-oauth-token.json"),
        )
        provider = google_sheets_adapter.UserOAuthTokenProvider(config)

        auth_url = provider._build_auth_url(
            redirect_uri="http://127.0.0.1:8765/oauth2callback",
            state="test-state",
            code_challenge="challenge-value",
        )

        self.assertIn("client_id=client-id", auth_url)
        self.assertIn("redirect_uri=http%3A%2F%2F127.0.0.1%3A8765%2Foauth2callback", auth_url)
        self.assertIn("code_challenge=challenge-value", auth_url)
        self.assertIn("code_challenge_method=S256", auth_url)
        self.assertIn("access_type=offline", auth_url)

    def test_user_oauth_token_provider_writes_cache_with_owner_only_permissions(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            token_path = pathlib.Path(tmpdir) / "token.json"
            config = google_sheets_adapter.UserOAuthConfig(
                spreadsheet_id="spreadsheet-123",
                client_id="client-id",
                client_secret=None,
                auth_uri="https://accounts.google.com/o/oauth2/v2/auth",
                token_uri="https://oauth2.googleapis.com/token",
                token_cache_path=token_path,
            )
            provider = google_sheets_adapter.UserOAuthTokenProvider(config)

            provider._write_cached_token(
                {
                    "access_token": "cached-token",
                    "refresh_token": "refresh-token",
                    "expires_at": time.time() + 3600,
                }
            )

            self.assertEqual(stat.S_IMODE(token_path.stat().st_mode), 0o600)

    def test_service_account_token_provider_writes_temp_key_with_owner_only_permissions(self):
        observed_modes = []

        def fake_run(args, input, capture_output, check):
            key_path = pathlib.Path(args[4])
            observed_modes.append(stat.S_IMODE(key_path.stat().st_mode))

            class Result:
                stdout = b"signed-bytes"

            return Result()

        config = google_sheets_adapter.ServiceAccountConfig(
            spreadsheet_id="spreadsheet-123",
            client_email="svc@example.com",
            private_key="-----BEGIN PRIVATE KEY-----\nFAKE\n-----END PRIVATE KEY-----\n",
        )
        provider = google_sheets_adapter.ServiceAccountTokenProvider(config)

        with mock.patch("google_sheets_adapter.subprocess.run", side_effect=fake_run):
            assertion = provider._build_jwt_assertion()

        self.assertTrue(assertion.endswith(".c2lnbmVkLWJ5dGVz"))
        self.assertEqual(observed_modes, [0o600])


class GoogleSheetsAdapterValueTests(unittest.TestCase):
    def test_dataframe_to_values_converts_headers_blanks_and_datetimes(self):
        df = pd.DataFrame(
            [
                {
                    "text": "alpha",
                    "count": 3,
                    "when": pd.Timestamp("2026-04-24T12:00:00Z"),
                    "empty": None,
                    "missing": pd.NA,
                }
            ]
        )

        values = google_sheets_adapter.dataframe_to_values(df)

        self.assertEqual(values[0], ["text", "count", "when", "empty", "missing"])
        self.assertEqual(values[1], ["alpha", 3, "2026-04-24T12:00:00+00:00", "", ""])


class GoogleSheetsAdapterWriteTests(unittest.TestCase):
    def test_write_tab_ensures_sheet_clears_and_updates_in_order(self):
        class RecordingAuthProvider:
            config = google_sheets_adapter.ServiceAccountConfig(
                spreadsheet_id="spreadsheet-123",
                client_email="svc@example.com",
                private_key="fake",
            )

            def get_access_token(self):
                return "token"

        class RecordingAdapter(google_sheets_adapter.GoogleSheetsAdapter):
            def __init__(self):
                super().__init__(RecordingAuthProvider())
                self.calls = []

            def ensure_sheet_exists(self, tab_name):
                self.calls.append(("ensure", tab_name))

            def clear_tab(self, tab_name):
                self.calls.append(("clear", tab_name))

            def update_tab_values(self, tab_name, rows):
                self.calls.append(("update", tab_name, rows))

        adapter = RecordingAdapter()
        df = pd.DataFrame([{"metric": "FRT", "value": 0.9}])

        adapter.write_tab("Refresh Control", df)

        self.assertEqual([call[0] for call in adapter.calls], ["ensure", "clear", "update"])
        self.assertEqual(adapter.calls[0][1], "Refresh Control")
        self.assertEqual(adapter.calls[2][1], "Refresh Control")
        self.assertEqual(adapter.calls[2][2], [["metric", "value"], ["FRT", 0.9]])


if __name__ == "__main__":
    unittest.main()
