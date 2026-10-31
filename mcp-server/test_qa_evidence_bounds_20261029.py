"""Regression coverage for bounded and JSON-safe QA evidence (#61)."""
import copy
import math
import unittest

import editing


def external(issues):
    return {"success": True,
            "passed": not any(issue["severity"] == "error" for issue in issues),
            "issues": issues}


class BoundedQaEvidenceTests(unittest.TestCase):
    def test_structural_overflow_stays_usable_and_fails_closed(self):
        items = [{"frame": 0, "length": 2, "layer": 0, "item_id": f"id-{i}"}
                 for i in range(1200)]
        report = editing.validate_timeline(items, duration=1)
        self.assertEqual(len(report["issues"]), 1000)
        self.assertEqual(report["issues"][-1]["code"], "QA_ISSUES_TRUNCATED")
        self.assertGreater(report["issues"][-1]["omitted_count"], 1000)
        self.assertEqual(report["issue_count"], 1000)
        self.assertFalse(report["passed"])
        self.assertEqual(editing.evaluate_qa_gate(report)["decision"], "repair")

    def test_exact_limit_does_not_add_truncation_marker(self):
        report = editing.validate_timeline(
            [], expected=[{"item_id": f"missing-{i}"} for i in range(1000)])
        self.assertEqual(report["issue_count"], 1000)
        self.assertEqual(report["issues"][-1]["code"], "EXPECTED_NOT_FOUND")
        self.assertEqual(editing.evaluate_qa_gate(report)["reason_code"], "QA_ISSUES_REMAIN")

    def test_composite_overflow_is_bounded_and_gate_compatible(self):
        checks = {
            "visual": external([{"code": "BLACK_FRAME", "severity": "warning"} for _ in range(700)]),
            "audio": external([{"code": "CLIPPING", "severity": "error"} for _ in range(700)]),
        }
        combined = editing.combine_qa_reports(
            editing.validate_timeline([]), checks,
            {"visual": {"frames": 20}, "audio": {"seconds": 20}})
        self.assertEqual(len(combined["issues"]), 1000)
        self.assertEqual(combined["issues"][-1]["code"], "QA_ISSUES_TRUNCATED")
        self.assertFalse(combined["passed"])
        self.assertEqual(editing.evaluate_qa_gate(combined)["decision"], "repair")

    def test_composite_warning_only_overflow_never_false_passes(self):
        issues = [{"code": "LOW_VOLUME", "severity": "warning"} for _ in range(1000)]
        result = editing.combine_qa_reports(
            editing.validate_timeline([], expected=[{"item_id": "absent"}]),
            {"audio": external(issues)}, {"audio": {}})
        self.assertFalse(result["passed"])
        self.assertEqual(result["issues"][-1]["severity"], "error")

    def test_below_limit_composite_preserves_all_evidence(self):
        issues = [{"code": "LOW_VOLUME", "severity": "warning"} for _ in range(1000)]
        result = editing.combine_qa_reports(editing.validate_timeline([]),
                                             {"audio": external(issues)}, {"audio": {}})
        self.assertEqual(result["issue_count"], 1000)
        self.assertTrue(result["passed"])
        self.assertEqual(result["issues"][-1]["code"], "LOW_VOLUME")

    def test_gate_rejects_non_finite_nested_issue_evidence(self):
        base = editing.validate_timeline([], expected=[{"item_id": "missing"}])
        for bad in (float("nan"), float("inf"), -float("inf"), {1, 2}, b"bad"):
            report = copy.deepcopy(base)
            report["issues"][0]["message"] = bad
            with self.subTest(bad=repr(bad)), self.assertRaisesRegex(ValueError, "JSON-safe"):
                editing.evaluate_qa_gate(report)

    def test_gate_rejects_non_json_history_evidence(self):
        base = editing.validate_timeline([], expected=[{"item_id": "missing"}])
        history = copy.deepcopy(base)
        history["issues"][0]["extra"] = {"bad": object()}
        with self.assertRaisesRegex(ValueError, "JSON-safe"):
            editing.evaluate_qa_gate(base, [history])

    def test_external_checks_reject_non_finite_issue_evidence(self):
        for bad in (float("nan"), float("inf"), {1, 2}, b"bad"):
            report = external([{"code": "BLACK_FRAME", "severity": "error", "detail": bad}])
            with self.subTest(bad=repr(bad)), self.assertRaisesRegex(ValueError, "JSON-safe"):
                editing.combine_qa_reports(editing.validate_timeline([]),
                                           {"visual": report}, {"visual": {}})

    def test_composite_criteria_reject_non_finite_values(self):
        for bad in (float("nan"), float("inf"), {1, 2}, b"bad"):
            with self.subTest(bad=repr(bad)), self.assertRaisesRegex(ValueError, "JSON-safe"):
                editing.combine_qa_reports(
                    editing.validate_timeline([]), {"visual": external([])},
                    {"visual": {"bad": bad}})

    def test_bounded_reports_are_stable_across_qa_history(self):
        items = [{"frame": 0, "length": 2, "layer": 0, "item_id": f"id-{i}"}
                 for i in range(1100)]
        before = editing.validate_timeline(items, duration=1)
        after = editing.validate_timeline(items, duration=1)
        self.assertEqual(editing.evaluate_qa_gate(after, [before])["reason_code"], "QA_STALLED")


if __name__ == "__main__":
    unittest.main()
