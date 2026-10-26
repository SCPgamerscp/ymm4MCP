"""Regression tests for trustworthy automated QA and duration repair hints."""
import copy
import unittest

import editing


class QaReportIntegrityTests(unittest.TestCase):
    def test_outside_duration_item_has_no_ineffective_single_edit(self):
        result = editing.validate_timeline([
            {"item_id": "late", "frame": 25, "length": 10, "layer": 1},
        ], duration=20)
        issue = next(i for i in result["issues"] if i["code"] == "EXCEEDS_DURATION")
        self.assertNotIn("suggested_fix", issue)
        self.assertIn("outside", issue["message"])
        self.assertFalse(result["passed"])

    def test_partly_visible_item_suggests_valid_length_trim(self):
        result = editing.validate_timeline([
            {"item_id": "partial", "frame": 15, "length": 10, "layer": 1},
        ], duration=20)
        issue = next(i for i in result["issues"] if i["code"] == "EXCEEDS_DURATION")
        self.assertEqual(issue["suggested_fix"]["prop"], "Length")
        self.assertEqual(issue["suggested_fix"]["value"], 5)

    def test_inconsistent_error_score_is_rejected(self):
        report = editing.validate_timeline([
            {"item_id": "late", "frame": 25, "length": 10, "layer": 1},
        ], duration=20)
        report["score"] = 100
        with self.assertRaisesRegex(ValueError, "score disagrees"):
            editing.evaluate_qa_gate(report)

    def test_inconsistent_warning_score_is_rejected(self):
        report = editing.validate_timeline([
            {"item_id": "first", "frame": 0, "length": 10, "layer": 1},
            {"item_id": "second", "frame": 20, "length": 10, "layer": 1},
        ])
        self.assertTrue(report["passed"])
        report["score"] = 100
        with self.assertRaisesRegex(ValueError, "score disagrees"):
            editing.evaluate_qa_gate(report)

    def test_invalid_item_ids_in_current_and_history_fail_cleanly(self):
        valid = editing.validate_timeline([
            {"item_id": "a", "frame": 0, "length": 20, "layer": 1},
            {"item_id": "b", "frame": 10, "length": 20, "layer": 1},
        ])
        for bad in (["a", 42], "a", [""], [None]):
            with self.subTest(item_ids=bad):
                invalid = copy.deepcopy(valid)
                invalid["issues"][0]["item_ids"] = bad
                with self.assertRaisesRegex(ValueError, "item_ids"):
                    editing.evaluate_qa_gate(invalid)
                with self.assertRaisesRegex(ValueError, "item_ids"):
                    editing.evaluate_qa_gate(valid, history=[invalid])

    def test_valid_reports_still_support_stall_detection(self):
        report = editing.validate_timeline([
            {"item_id": "a", "frame": 0, "length": 20, "layer": 1},
            {"item_id": "b", "frame": 10, "length": 20, "layer": 1},
        ])
        self.assertEqual(editing.evaluate_qa_gate(report, [copy.deepcopy(report)])["reason_code"],
                         "QA_STALLED")


if __name__ == "__main__":
    unittest.main()
