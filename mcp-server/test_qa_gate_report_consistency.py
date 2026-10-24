"""Reject contradictory or malformed QA history before repair decisions."""

import copy
import unittest

import editing


class QaGateReportConsistencyTests(unittest.TestCase):
    def setUp(self):
        self.clean = editing.validate_timeline([])

    def test_error_issue_cannot_be_reported_as_passed(self):
        inconsistent = copy.deepcopy(self.clean)
        inconsistent["issues"].append({"code": "AUDIO_CLIPPING", "severity": "error"})
        with self.assertRaisesRegex(ValueError, "passed status disagrees"):
            editing.evaluate_qa_gate(inconsistent)

    def test_history_cannot_hide_an_error_issue(self):
        inconsistent = copy.deepcopy(self.clean)
        inconsistent["issues"].append({"code": "AUDIO_CLIPPING", "severity": "error"})
        with self.assertRaisesRegex(ValueError, "passed status disagrees"):
            editing.evaluate_qa_gate(self.clean, history=[inconsistent])

    def test_failed_report_without_errors_is_rejected(self):
        inconsistent = copy.deepcopy(self.clean)
        inconsistent["passed"] = False
        with self.assertRaisesRegex(ValueError, "passed status disagrees"):
            editing.evaluate_qa_gate(inconsistent)

    def test_unknown_issue_severity_is_rejected(self):
        invalid = copy.deepcopy(self.clean)
        invalid["issues"].append({"code": "UNKNOWN", "severity": "fatal"})
        with self.assertRaisesRegex(ValueError, "error or warning severity"):
            editing.evaluate_qa_gate(invalid)

    def test_warning_only_report_can_pass(self):
        items = [{"frame": 0, "length": 10, "layer": 1},
                 {"frame": 20, "length": 10, "layer": 1}]
        warning = editing.validate_timeline(items)
        self.assertTrue(warning["passed"])
        self.assertEqual(editing.evaluate_qa_gate(warning)["decision"], "pass")

    def test_genuine_error_report_still_requests_repair(self):
        items = [{"frame": 0, "length": 10, "layer": 1},
                 {"frame": 5, "length": 10, "layer": 1}]
        failed = editing.validate_timeline(items)
        self.assertFalse(failed["passed"])
        self.assertEqual(editing.evaluate_qa_gate(failed)["decision"], "repair")


if __name__ == "__main__":
    unittest.main()
