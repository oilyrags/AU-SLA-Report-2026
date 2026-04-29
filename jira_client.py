#!/usr/bin/env python3
from __future__ import annotations

import base64
import getpass
import os
import sys
import time
from dataclasses import dataclass
from typing import Any, Dict, List, Optional

import requests

from sla_rules import DEFAULT_TIMEOUT, RETRY_STATUS_CODES, SEARCH_PAGE_SIZE


@dataclass
class JiraConfig:
    base_url: str
    email: str
    api_token: str
    timeout: int = DEFAULT_TIMEOUT

    def __post_init__(self) -> None:
        self.base_url = self.base_url.strip().strip("\"'")


class JiraClient:
    def __init__(self, config: JiraConfig) -> None:
        self.config = config
        self.session = requests.Session()
        raw = f"{config.email}:{config.api_token}".encode("utf-8")
        auth = base64.b64encode(raw).decode("ascii")
        self.session.headers.update(
            {
                "Authorization": f"Basic {auth}",
                "Accept": "application/json",
                "Content-Type": "application/json",
                "User-Agent": "smg-sla-report/1.0",
            }
        )

    def _request(
        self,
        method: str,
        path: str,
        *,
        params: Optional[Dict[str, Any]] = None,
        json_body: Optional[Dict[str, Any]] = None,
        timeout: Optional[int] = None,
        max_retries: int = 5,
    ) -> requests.Response:
        url = f"{self.config.base_url.rstrip('/')}{path}"

        for attempt in range(max_retries + 1):
            resp = self.session.request(
                method=method,
                url=url,
                params=params,
                json=json_body,
                timeout=timeout or self.config.timeout,
            )
            if resp.status_code not in RETRY_STATUS_CODES:
                return resp
            if attempt == max_retries:
                return resp

            retry_after = resp.headers.get("Retry-After")
            sleep_s = int(retry_after) if retry_after and retry_after.isdigit() else min(2 ** attempt, 30)
            print(f"[warn] {method} {path} -> {resp.status_code}; retry in {sleep_s}s", file=sys.stderr)
            time.sleep(sleep_s)

        return resp

    @staticmethod
    def raise_for_status(resp: requests.Response, context: str) -> None:
        if resp.ok:
            return
        try:
            detail = resp.json()
        except Exception:
            detail = resp.text
        raise RuntimeError(f"{context} failed: HTTP {resp.status_code} - {detail}")

    def get_server_info(self) -> Dict[str, Any]:
        resp = self._request("GET", "/rest/api/3/serverInfo")
        self.raise_for_status(resp, "server info")
        return resp.json()

    def search_issues(self, jql: str, fields: List[str], page_size: int = SEARCH_PAGE_SIZE) -> List[Dict[str, Any]]:
        issues: List[Dict[str, Any]] = []
        next_page_token = None

        while True:
            body = {
                "jql": jql,
                "maxResults": page_size,
                "fields": fields,
                "fieldsByKeys": False,
            }
            if next_page_token:
                body["nextPageToken"] = next_page_token

            resp = self._request("POST", "/rest/api/3/search/jql", json_body=body)
            self.raise_for_status(resp, "search issues")
            payload = resp.json()

            batch = payload.get("issues", []) or []
            issues.extend(batch)

            next_page_token = payload.get("nextPageToken")
            is_last = payload.get("isLast", next_page_token is None)

            print(f"[info] fetched {len(issues)} issues", file=sys.stderr)

            if is_last or not batch:
                break

        return issues

    def get_issue_changelog(self, issue_key: str) -> Dict[str, Any]:
        resp = self._request("GET", f"/rest/api/3/issue/{issue_key}/changelog")
        self.raise_for_status(resp, f"changelog {issue_key}")
        return resp.json()

    def get_sla_cycles(self, issue_key: str) -> Dict[str, Any]:
        resp = self._request("GET", f"/rest/servicedeskapi/request/{issue_key}/sla")
        self.raise_for_status(resp, f"SLA {issue_key}")
        return resp.json()


def prompt_env(name: str, label: str, secret: bool = False) -> str:
    val = os.environ.get(name)
    if val:
        return val
    return getpass.getpass(f"{label}: ").strip() if secret else input(f"{label}: ").strip()
