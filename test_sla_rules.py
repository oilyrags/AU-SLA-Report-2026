#!/usr/bin/env python3
import os
import unittest
from unittest import mock

import sla_rules


class SlaRulesTests(unittest.TestCase):
    def test_report_scope_uses_defaults_env_and_explicit_values(self):
        with mock.patch.dict(os.environ, {}, clear=True):
            self.assertEqual(
                sla_rules.resolve_report_scope(project_key=None, year=None),
                ("ASD", 2026),
            )

        with mock.patch.dict(os.environ, {"JIRA_PROJECT_KEY": "OPS", "SLA_REPORT_YEAR": "2027"}, clear=True):
            self.assertEqual(
                sla_rules.resolve_report_scope(project_key=None, year=None),
                ("OPS", 2027),
            )
            self.assertEqual(
                sla_rules.resolve_report_scope(project_key="ASD", year=2026),
                ("ASD", 2026),
            )

    def test_report_scope_rejects_invalid_year(self):
        with self.assertRaisesRegex(ValueError, "SLA_REPORT_YEAR"):
            sla_rules.resolve_report_scope(project_key=None, year="twenty")

    def test_targets_and_aliases_are_shared(self):
        self.assertEqual(sla_rules.normalize_issue_type(" emal   request "), "email request")
        self.assertEqual(sla_rules.get_targets("P2 - High", "Q2"), (1.0, 0.95))
        self.assertEqual(sla_rules.get_targets("P2 – High", "Q2"), (1.0, 0.95))


if __name__ == "__main__":
    unittest.main()
