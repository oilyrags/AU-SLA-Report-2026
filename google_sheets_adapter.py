#!/usr/bin/env python3
from __future__ import annotations

import base64
import hashlib
import json
import os
import secrets
import socket
import subprocess
import tempfile
import time
import webbrowser
from dataclasses import dataclass
from datetime import datetime
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from typing import Any, Dict, List, Optional
from urllib.parse import parse_qs, quote, urlencode, urlparse

import pandas as pd
import requests


DEFAULT_TOKEN_URI = "https://oauth2.googleapis.com/token"
DEFAULT_AUTH_URI = "https://accounts.google.com/o/oauth2/v2/auth"
SHEETS_SCOPE = "https://www.googleapis.com/auth/spreadsheets"
DEFAULT_TOKEN_CACHE_PATH = Path.home() / ".config" / "sla-report" / "google-oauth-token.json"


@dataclass(frozen=True)
class ServiceAccountConfig:
    spreadsheet_id: str
    client_email: str
    private_key: str
    token_uri: str = DEFAULT_TOKEN_URI
    scopes: tuple[str, ...] = (SHEETS_SCOPE,)


@dataclass(frozen=True)
class UserOAuthConfig:
    spreadsheet_id: str
    client_id: str
    client_secret: Optional[str]
    auth_uri: str
    token_uri: str
    token_cache_path: Path
    scopes: tuple[str, ...] = (SHEETS_SCOPE,)


def _load_json_payload_from_path_or_env(
    *,
    path_key: str,
    raw_key: str,
    env: Dict[str, str],
) -> Optional[Dict[str, Any]]:
    json_path = env.get(path_key)
    if json_path:
        with Path(json_path).expanduser().open("r", encoding="utf-8") as f:
            return json.load(f)

    raw_json = env.get(raw_key)
    if raw_json:
        return json.loads(raw_json)

    return None


def _load_service_account_payload(env: Dict[str, str]) -> Optional[Dict[str, Any]]:
    return _load_json_payload_from_path_or_env(
        path_key="GOOGLE_SERVICE_ACCOUNT_JSON_PATH",
        raw_key="GOOGLE_SERVICE_ACCOUNT_JSON",
        env=env,
    )


def _load_oauth_client_payload(env: Dict[str, str]) -> Optional[Dict[str, Any]]:
    return _load_json_payload_from_path_or_env(
        path_key="GOOGLE_OAUTH_CLIENT_SECRET_JSON_PATH",
        raw_key="GOOGLE_OAUTH_CLIENT_SECRET_JSON",
        env=env,
    )


def load_service_account_config(env: Optional[Dict[str, str]] = None) -> Optional[ServiceAccountConfig]:
    env = os.environ if env is None else env
    spreadsheet_id = env.get("GOOGLE_SHEETS_SPREADSHEET_ID")
    if not spreadsheet_id:
        return None

    payload = _load_service_account_payload(env)
    if not payload:
        return None

    client_email = payload.get("client_email")
    private_key = payload.get("private_key")
    token_uri = payload.get("token_uri", DEFAULT_TOKEN_URI)

    if not client_email or not private_key:
        raise ValueError("Google service account JSON is missing client_email or private_key")

    return ServiceAccountConfig(
        spreadsheet_id=spreadsheet_id,
        client_email=client_email,
        private_key=private_key,
        token_uri=token_uri,
    )


def load_user_oauth_config(env: Optional[Dict[str, str]] = None) -> Optional[UserOAuthConfig]:
    env = os.environ if env is None else env
    spreadsheet_id = env.get("GOOGLE_SHEETS_SPREADSHEET_ID")
    if not spreadsheet_id:
        return None

    payload = _load_oauth_client_payload(env)
    if not payload:
        return None

    client_payload = payload.get("installed") or payload.get("web") or payload
    client_id = client_payload.get("client_id")
    if not client_id:
        raise ValueError("Google OAuth client JSON is missing client_id")

    token_cache_path = Path(
        env.get("GOOGLE_OAUTH_TOKEN_PATH", str(DEFAULT_TOKEN_CACHE_PATH))
    ).expanduser()

    return UserOAuthConfig(
        spreadsheet_id=spreadsheet_id,
        client_id=client_id,
        client_secret=client_payload.get("client_secret"),
        auth_uri=client_payload.get("auth_uri", DEFAULT_AUTH_URI),
        token_uri=client_payload.get("token_uri", DEFAULT_TOKEN_URI),
        token_cache_path=token_cache_path,
    )


def load_google_sheets_adapter_from_env(env: Optional[Dict[str, str]] = None) -> Optional["GoogleSheetsAdapter"]:
    env = os.environ if env is None else env

    oauth_config = load_user_oauth_config(env=env)
    if oauth_config is not None:
        return GoogleSheetsAdapter(UserOAuthTokenProvider(oauth_config))

    service_account_config = load_service_account_config(env=env)
    if service_account_config is not None:
        return GoogleSheetsAdapter(ServiceAccountTokenProvider(service_account_config))

    return None


def dataframe_to_values(df: pd.DataFrame) -> List[List[Any]]:
    rows: List[List[Any]] = [list(df.columns)]

    for _, series in df.iterrows():
        rows.append([_convert_cell_value(value) for value in series.tolist()])

    return rows


def _convert_cell_value(value: Any) -> Any:
    if value is None:
        return ""

    if isinstance(value, float) and pd.isna(value):
        return ""

    if value is pd.NA:
        return ""

    if isinstance(value, pd.Timestamp):
        if pd.isna(value):
            return ""
        return value.to_pydatetime().isoformat()

    if isinstance(value, datetime):
        return value.isoformat()

    if hasattr(value, "item") and not isinstance(value, (str, bytes, bytearray, dict, list, tuple)):
        try:
            value = value.item()
        except Exception:
            pass

    if isinstance(value, bool):
        return value

    if isinstance(value, (int, float, str)):
        if isinstance(value, float) and pd.isna(value):
            return ""
        return value

    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass

    return str(value)


def _base64url(data: bytes) -> str:
    return base64.urlsafe_b64encode(data).rstrip(b"=").decode("ascii")


def _json_segment(payload: Dict[str, Any]) -> str:
    return _base64url(json.dumps(payload, separators=(",", ":"), sort_keys=True).encode("utf-8"))


def _quoted_sheet_range(tab_name: str, cell: str = "A1") -> str:
    escaped = tab_name.replace("'", "''")
    if " " in tab_name or any(ch in tab_name for ch in ("!", ":", "/", "?")):
        return f"'{escaped}'!{cell}"
    return f"{escaped}!{cell}"


def _pkce_verifier() -> str:
    return secrets.token_urlsafe(64)


def _pkce_challenge(verifier: str) -> str:
    return _base64url(hashlib.sha256(verifier.encode("ascii")).digest())


def _pick_free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        sock.listen(1)
        return int(sock.getsockname()[1])


class OAuthCallbackServer(HTTPServer):
    def __init__(self, server_address):
        super().__init__(server_address, OAuthCallbackHandler)
        self.authorization_code: Optional[str] = None
        self.state: Optional[str] = None
        self.error: Optional[str] = None


class OAuthCallbackHandler(BaseHTTPRequestHandler):
    def do_GET(self):  # noqa: N802
        parsed = urlparse(self.path)
        query = parse_qs(parsed.query)
        self.server.authorization_code = query.get("code", [None])[0]
        self.server.state = query.get("state", [None])[0]
        self.server.error = query.get("error", [None])[0]

        if self.server.authorization_code and not self.server.error:
            body = b"Authentication complete. You can close this window."
            self.send_response(200)
        else:
            body = b"Authentication failed. Return to Terminal for details."
            self.send_response(400)

        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, format, *args):  # noqa: A003
        return


class ServiceAccountTokenProvider:
    def __init__(
        self,
        config: ServiceAccountConfig,
        *,
        session: Optional[requests.Session] = None,
        timeout_seconds: int = 60,
    ) -> None:
        self.config = config
        self.session = session or requests.Session()
        self.timeout_seconds = timeout_seconds
        self._access_token: Optional[str] = None

    def get_access_token(self) -> str:
        if self._access_token:
            return self._access_token

        assertion = self._build_jwt_assertion()
        resp = self.session.post(
            self.config.token_uri,
            data={
                "grant_type": "urn:ietf:params:oauth:grant-type:jwt-bearer",
                "assertion": assertion,
            },
            timeout=self.timeout_seconds,
            headers={"Content-Type": "application/x-www-form-urlencoded"},
        )
        if not resp.ok:
            raise RuntimeError(f"Google token request failed: HTTP {resp.status_code} - {resp.text}")

        payload = resp.json()
        token = payload.get("access_token")
        if not token:
            raise RuntimeError("Google token response did not include access_token")

        self._access_token = token
        return token

    def _build_jwt_assertion(self) -> str:
        now = int(time.time())
        claims = {
            "iss": self.config.client_email,
            "scope": " ".join(self.config.scopes),
            "aud": self.config.token_uri,
            "iat": now,
            "exp": now + 3600,
        }
        header = {"alg": "RS256", "typ": "JWT"}
        signing_input = f"{_json_segment(header)}.{_json_segment(claims)}".encode("ascii")

        fd, key_path = tempfile.mkstemp(suffix=".pem")
        os.chmod(key_path, 0o600)
        with os.fdopen(fd, "w", encoding="utf-8") as tmp_key:
            tmp_key.write(self.config.private_key)

        try:
            signed = subprocess.run(
                ["openssl", "dgst", "-sha256", "-sign", key_path, "-binary"],
                input=signing_input,
                capture_output=True,
                check=True,
            ).stdout
        finally:
            try:
                os.unlink(key_path)
            except FileNotFoundError:
                pass

        return f"{signing_input.decode('ascii')}.{_base64url(signed)}"


class UserOAuthTokenProvider:
    def __init__(
        self,
        config: UserOAuthConfig,
        *,
        session: Optional[requests.Session] = None,
        timeout_seconds: int = 60,
    ) -> None:
        self.config = config
        self.session = session or requests.Session()
        self.timeout_seconds = timeout_seconds

    def get_access_token(self) -> str:
        token_payload = self._load_cached_token()
        if token_payload and self._token_is_valid(token_payload):
            return token_payload["access_token"]

        if token_payload and token_payload.get("refresh_token"):
            refreshed = self._refresh_access_token(token_payload["refresh_token"])
            self._write_cached_token(refreshed)
            return refreshed["access_token"]

        created = self._authorize_with_browser()
        self._write_cached_token(created)
        return created["access_token"]

    def _load_cached_token(self) -> Optional[Dict[str, Any]]:
        if not self.config.token_cache_path.exists():
            return None
        with self.config.token_cache_path.open("r", encoding="utf-8") as f:
            return json.load(f)

    def _write_cached_token(self, token_payload: Dict[str, Any]) -> None:
        self.config.token_cache_path.parent.mkdir(mode=0o700, parents=True, exist_ok=True)
        fd = os.open(self.config.token_cache_path, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(token_payload, f, ensure_ascii=False, indent=2)
        os.chmod(self.config.token_cache_path, 0o600)

    def _token_is_valid(self, token_payload: Dict[str, Any]) -> bool:
        access_token = token_payload.get("access_token")
        expires_at = token_payload.get("expires_at")
        if not access_token or not expires_at:
            return False
        return float(expires_at) > time.time() + 60

    def _refresh_access_token(self, refresh_token: str) -> Dict[str, Any]:
        data = {
            "grant_type": "refresh_token",
            "refresh_token": refresh_token,
            "client_id": self.config.client_id,
        }
        if self.config.client_secret:
            data["client_secret"] = self.config.client_secret

        resp = self.session.post(
            self.config.token_uri,
            data=data,
            timeout=self.timeout_seconds,
            headers={"Content-Type": "application/x-www-form-urlencoded"},
        )
        if not resp.ok:
            raise RuntimeError(f"Google OAuth token refresh failed: HTTP {resp.status_code} - {resp.text}")

        payload = resp.json()
        access_token = payload.get("access_token")
        expires_in = payload.get("expires_in")
        if not access_token or not expires_in:
            raise RuntimeError("Google OAuth refresh response missing access_token or expires_in")

        return {
            "access_token": access_token,
            "refresh_token": payload.get("refresh_token", refresh_token),
            "expires_at": time.time() + float(expires_in),
            "token_type": payload.get("token_type", "Bearer"),
        }

    def _authorize_with_browser(self) -> Dict[str, Any]:
        verifier = _pkce_verifier()
        challenge = _pkce_challenge(verifier)
        state = secrets.token_urlsafe(24)
        port = _pick_free_port()
        redirect_uri = f"http://127.0.0.1:{port}/oauth2callback"
        auth_url = self._build_auth_url(
            redirect_uri=redirect_uri,
            state=state,
            code_challenge=challenge,
        )

        print("Open this URL to authorize Google Sheets access:", auth_url)
        try:
            webbrowser.open(auth_url)
        except Exception:
            pass

        server = OAuthCallbackServer(("127.0.0.1", port))
        server.timeout = 300
        server.handle_request()

        if server.error:
            raise RuntimeError(f"Google OAuth authorization failed: {server.error}")
        if not server.authorization_code:
            raise RuntimeError("Timed out waiting for Google OAuth authorization callback")
        if server.state != state:
            raise RuntimeError("Google OAuth state mismatch during callback")

        data = {
            "grant_type": "authorization_code",
            "code": server.authorization_code,
            "client_id": self.config.client_id,
            "redirect_uri": redirect_uri,
            "code_verifier": verifier,
        }
        if self.config.client_secret:
            data["client_secret"] = self.config.client_secret

        resp = self.session.post(
            self.config.token_uri,
            data=data,
            timeout=self.timeout_seconds,
            headers={"Content-Type": "application/x-www-form-urlencoded"},
        )
        if not resp.ok:
            raise RuntimeError(f"Google OAuth token exchange failed: HTTP {resp.status_code} - {resp.text}")

        payload = resp.json()
        access_token = payload.get("access_token")
        expires_in = payload.get("expires_in")
        refresh_token = payload.get("refresh_token")
        if not access_token or not expires_in:
            raise RuntimeError("Google OAuth token exchange missing access_token or expires_in")
        if not refresh_token:
            raise RuntimeError("Google OAuth token exchange did not return refresh_token; try revoking prior consent and retrying")

        return {
            "access_token": access_token,
            "refresh_token": refresh_token,
            "expires_at": time.time() + float(expires_in),
            "token_type": payload.get("token_type", "Bearer"),
        }

    def _build_auth_url(self, *, redirect_uri: str, state: str, code_challenge: str) -> str:
        params = {
            "client_id": self.config.client_id,
            "redirect_uri": redirect_uri,
            "response_type": "code",
            "scope": " ".join(self.config.scopes),
            "access_type": "offline",
            "prompt": "consent",
            "state": state,
            "code_challenge": code_challenge,
            "code_challenge_method": "S256",
        }
        return f"{self.config.auth_uri}?{urlencode(params)}"


class GoogleSheetsAdapter:
    def __init__(
        self,
        auth_provider,
        *,
        session: Optional[requests.Session] = None,
        timeout_seconds: int = 60,
    ) -> None:
        self.auth_provider = auth_provider
        self.config = auth_provider.config
        self.session = session or requests.Session()
        self.timeout_seconds = timeout_seconds

    def write_tab(self, tab_name: str, dataframe: pd.DataFrame) -> None:
        rows = dataframe_to_values(dataframe)
        self.ensure_sheet_exists(tab_name)
        self.clear_tab(tab_name)
        self.update_tab_values(tab_name, rows)

    def ensure_sheet_exists(self, tab_name: str) -> None:
        metadata = self._request_json(
            "GET",
            f"/v4/spreadsheets/{self.config.spreadsheet_id}",
            params={"fields": "sheets.properties.title"},
        )
        existing = {
            sheet.get("properties", {}).get("title")
            for sheet in metadata.get("sheets", [])
        }
        if tab_name in existing:
            return

        body = {"requests": [{"addSheet": {"properties": {"title": tab_name}}}]}
        self._request_json(
            "POST",
            f"/v4/spreadsheets/{self.config.spreadsheet_id}:batchUpdate",
            json_body=body,
        )

    def clear_tab(self, tab_name: str) -> None:
        range_name = quote(_quoted_sheet_range(tab_name, "A1"), safe="")
        self._request_json(
            "POST",
            f"/v4/spreadsheets/{self.config.spreadsheet_id}/values/{range_name}:clear",
            json_body={},
        )

    def update_tab_values(self, tab_name: str, rows: List[List[Any]]) -> None:
        range_name = quote(_quoted_sheet_range(tab_name, "A1"), safe="")
        self._request_json(
            "PUT",
            f"/v4/spreadsheets/{self.config.spreadsheet_id}/values/{range_name}",
            params={"valueInputOption": "RAW"},
            json_body={"majorDimension": "ROWS", "values": rows},
        )

    def read_tab_values(self, tab_name: str) -> List[List[Any]]:
        range_name = quote(_quoted_sheet_range(tab_name, "A1"), safe="")
        payload = self._request_json(
            "GET",
            f"/v4/spreadsheets/{self.config.spreadsheet_id}/values/{range_name}",
        )
        return payload.get("values", [])

    def _request_json(
        self,
        method: str,
        path: str,
        *,
        params: Optional[Dict[str, Any]] = None,
        json_body: Optional[Dict[str, Any]] = None,
    ) -> Dict[str, Any]:
        url = f"https://sheets.googleapis.com{path}"
        headers = {
            "Authorization": f"Bearer {self.auth_provider.get_access_token()}",
            "Accept": "application/json",
            "Content-Type": "application/json",
        }
        resp = self.session.request(
            method=method,
            url=url,
            params=params,
            json=json_body,
            headers=headers,
            timeout=self.timeout_seconds,
        )
        if not resp.ok:
            try:
                detail = resp.json()
            except Exception:
                detail = resp.text
            raise RuntimeError(f"Sheets API {method} {path} failed: HTTP {resp.status_code} - {detail}")

        if resp.status_code == 204:
            return {}
        return resp.json()
