#!/usr/bin/env python3
import unittest

from jira_client import JiraConfig


class JiraConfigTests(unittest.TestCase):
    def test_base_url_strips_accidental_wrapping_quotes(self):
        config = JiraConfig(
            base_url='"https://smg-au.atlassian.net"',
            email="user@example.com",
            api_token="token",
        )

        self.assertEqual(config.base_url, "https://smg-au.atlassian.net")


if __name__ == "__main__":
    unittest.main()
