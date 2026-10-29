"""Regression tests for bounded QA progress and complete evidence signatures (#61)."""
import copy
import unittest

import editing


def overflowing_report(count):
    # Separate layers avoid overlap noise; the first 999 issues are identical.
    items = [{"frame": 0, "length": 2, "layer": i, "item_id": f"id-{i}"}
             for i in range(count)]
    return editing.validate_timeline(items, duration=1)


class QaEvidenceProgressTests(unittest.TestCase):
    def test_more_omitted_issues_at_score_floor_is_regression(self):
        before = overflowing_report(1100)
        after = overflowing_report(1200)
        self.assertEqual(before["score"], after["score"])
        self.assertEqual(before["summary"], after["summary"])
        self.assertGreater(after["issues"][-1]["omitted_count"],
                           before["issues"][-1]["omitted_count"])
        result = editing.evaluate_qa_gate(after, [before])
        self.assertEqual(result["reason_code"], "QA_REGRESSED")
        self.assertEqual(result["decision"], "stop")
        self.assertEqual(result["suggested_action"], "consider_checkpoint_rollback")

    def test_fewer_omitted_issues_is_progress_not_stall(self):
        before = overflowing_report(1200)
        after = overflowing_report(1100)
        self.assertEqual(editing.evaluate_qa_gate(after, [before])["reason_code"],
                         "QA_ISSUES_REMAIN")

    def test_malformed_truncation_marker_is_rejected(self):
        base = overflowing_report(1100)
        for bad in (None, 0, -1, True, "101", 1.5):
            report = copy.deepcopy(base)
            report["issues"][-1]["omitted_count"] = bad
            with self.subTest(bad=repr(bad)), self.assertRaisesRegex(ValueError, "omitted_count"):
                editing.evaluate_qa_gate(report)

    def test_truncation_marker_must_be_error(self):
        report = overflowing_report(1100)
        report["issues"][-1]["severity"] = "warning"
        report["summary"] = {"errors": 999, "warnings": 1}
        with self.assertRaisesRegex(ValueError, "omitted_count"):
            editing.evaluate_qa_gate(report)

    def test_external_forged_truncation_marker_rejected_during_composition(self):
        for omitted in (None, 0, -3, True, "3"):
            issue = {"code": "QA_ISSUES_TRUNCATED", "severity": "error",
                     "omitted_count": omitted}
            with self.subTest(omitted=repr(omitted)), self.assertRaisesRegex(
                    ValueError, "truncation marker"):
                editing.combine_qa_reports(editing.validate_timeline([]),
                                           {"visual": {"success": True, "passed": False,
                                                       "issues": [issue]}},
                                           {"visual": {}})

    def test_valid_external_truncation_marker_remains_supported(self):
        issue = {"code": "QA_ISSUES_TRUNCATED", "severity": "error",
                 "omitted_count": 8}
        result = editing.combine_qa_reports(
            editing.validate_timeline([]),
            {"visual": {"success": True, "passed": False, "issues": [issue]}},
            {"visual": {}})
        self.assertFalse(result["passed"])
        self.assertEqual(result["issues"][0]["omitted_count"], 8)
        self.assertEqual(editing.evaluate_qa_gate(result)["decision"], "repair")

    def test_equal_truncation_evidence_stalls(self):
        before = overflowing_report(1100)
        after = overflowing_report(1100)
        self.assertEqual(editing.evaluate_qa_gate(after, [before])["reason_code"],
                         "QA_STALLED")

    def test_new_external_issue_detail_is_progress(self):
        before = editing.validate_timeline([], expected=[{"item_id": "missing"}])
        after = copy.deepcopy(before)
        after["issues"][0]["diagnostics"] = {"observed": "new evidence"}
        self.assertEqual(editing.evaluate_qa_gate(after, [before])["reason_code"],
                         "QA_ISSUES_REMAIN")

    def test_stable_ids_are_order_insensitive(self):
        before = editing.validate_timeline([
            {"frame": 0, "length": 4, "layer": 0, "item_id": "a"},
            {"frame": 1, "length": 4, "layer": 0, "item_id": "b"},
        ])
        after = copy.deepcopy(before)
        overlap = next(i for i in after["issues"] if i["code"] == "OVERLAP")
        overlap["item_ids"].reverse()
        overlap["indices"].reverse()
        self.assertEqual(editing.evaluate_qa_gate(after, [before])["reason_code"],
                         "QA_STALLED")

    def test_positional_index_only_change_is_not_progress(self):
        before = editing.validate_timeline([], expected=[{"item_id": "missing"}])
        after = copy.deepcopy(before)
        after["issues"][0]["index"] = 200
        self.assertEqual(editing.evaluate_qa_gate(after, [before])["reason_code"],
                         "QA_STALLED")


if __name__ == "__main__":
    unittest.main()
