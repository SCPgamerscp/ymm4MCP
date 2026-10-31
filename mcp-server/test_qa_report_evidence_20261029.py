"""Regression coverage for composite QA evidence and safe stall decisions (#61)."""
import copy
import unittest

import editing


def structural_failure():
    return editing.validate_timeline([], expected=[{"item_id": "missing"}])


def external_report(issues=None, **metadata):
    if issues is None:
        issues = [{"code": "BLACK_FRAME", "severity": "error"}]
    return {"success": True, "passed": not any(x["severity"] == "error" for x in issues),
            "issues": issues, **metadata}


def combined(report):
    return editing.combine_qa_reports(editing.validate_timeline([]),
                                      {"visual": report}, {"visual": {"end_frame": 30}})


class CompositeEvidenceIntegrityTests(unittest.TestCase):
    def test_rejects_external_valid_flag_disagreement(self):
        for valid in (True, None, 1):
            with self.subTest(valid=valid), self.assertRaisesRegex(ValueError, "valid status disagrees"):
                combined(external_report(valid=valid))

    def test_rejects_external_issue_count_mismatch_and_boolean(self):
        for count in (0, True, "1", -1):
            with self.subTest(count=count), self.assertRaisesRegex(ValueError, "issue_count disagrees"):
                combined(external_report(issue_count=count))

    def test_rejects_external_summary_mismatch_and_boolean(self):
        for summary in ({"errors": 0, "warnings": 0},
                        {"errors": True, "warnings": 0},
                        {"errors": 1, "warnings": False},
                        {"errors": 1},
                        None):
            with self.subTest(summary=summary), self.assertRaisesRegex(ValueError, "summary disagrees"):
                combined(external_report(summary=summary))

    def test_rejects_external_problems_alias_disagreement(self):
        with self.assertRaisesRegex(ValueError, "problems disagrees"):
            combined(external_report(problems=[]))

    def test_accepts_consistent_external_metadata(self):
        issue = {"code": "BLACK_FRAME", "severity": "error"}
        result = combined(external_report([issue], valid=False, issue_count=1,
                                          summary={"errors": 1, "warnings": 0},
                                          problems=[issue]))
        self.assertFalse(result["passed"])
        self.assertEqual(result["summary"], {"errors": 1, "warnings": 0})

    def test_accepts_optional_metadata_absence_for_legacy_checks(self):
        self.assertFalse(combined(external_report())["passed"])


class GateEvidenceIntegrityTests(unittest.TestCase):
    def test_rejects_oversized_qa_issue_history(self):
        report = structural_failure()
        oversized = {**report, "issues": report["issues"] * 1001}
        with self.assertRaisesRegex(ValueError, "at most 1000 issues"):
            editing.evaluate_qa_gate(oversized)

    def test_rejects_inconsistent_structural_problems_alias(self):
        report = structural_failure()
        with self.assertRaisesRegex(ValueError, "problems disagrees"):
            editing.evaluate_qa_gate({**report, "problems": []})

    def test_different_expected_text_is_a_new_error(self):
        before = structural_failure()
        after = copy.deepcopy(before)
        before["issues"][0]["expected_text"] = "A"
        after["issues"][0]["expected_text"] = "B"
        self.assertEqual(editing.evaluate_qa_gate(after, [before])["reason_code"],
                         "QA_REGRESSED")

    def test_different_actual_text_does_not_look_stalled(self):
        before = structural_failure()
        after = copy.deepcopy(before)
        before["issues"][0]["actual_texts"] = ["A"]
        after["issues"][0]["actual_texts"] = ["B"]
        self.assertEqual(editing.evaluate_qa_gate(after, [before])["reason_code"],
                         "QA_ISSUES_REMAIN")

    def test_changed_repair_hint_does_not_look_stalled(self):
        before = structural_failure()
        after = copy.deepcopy(before)
        before["issues"][0]["suggested_fix"] = {"action": "edit_item", "value": 10}
        after["issues"][0]["suggested_fix"] = {"action": "edit_item", "value": 20}
        self.assertEqual(editing.evaluate_qa_gate(after, [before])["reason_code"],
                         "QA_ISSUES_REMAIN")

    def test_unchanged_issue_still_stalls(self):
        report = structural_failure()
        self.assertEqual(editing.evaluate_qa_gate(report, [copy.deepcopy(report)])["reason_code"],
                         "QA_STALLED")


if __name__ == "__main__":
    unittest.main()
