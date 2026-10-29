"""Regression coverage for QA trust boundaries (#61)."""
import unittest
import editing


def external(issues):
    return {"success": True, "passed": not any(i["severity"] == "error" for i in issues),
            "issues": issues}


class QaTrustBoundaryTests(unittest.TestCase):
    def setUp(self):
        self.structural = editing.validate_timeline([])

    def combine(self, issues, source="visual", criteria=None):
        return editing.combine_qa_reports(
            self.structural, {source: external(issues)},
            {source: {} if criteria is None else criteria})

    def test_extreme_elapsed_integer_is_value_error(self):
        report = editing.validate_timeline([], expected=[{"item_id": "missing"}])
        with self.assertRaisesRegex(ValueError, "elapsed_seconds"):
            editing.evaluate_qa_gate(report, elapsed_seconds=10 ** 1000)

    def test_extreme_script_speed_is_value_error(self):
        with self.assertRaisesRegex(ValueError, "chars_per_sec"):
            editing.plan_script({"lines": [{"text": "hi", "character": "a"}],
                                 "chars_per_sec": 10 ** 1000})

    def test_unpaired_surrogate_in_qa_evidence(self):
        report = editing.validate_timeline([], expected=[{"item_id": "missing"}])
        report["issues"][0]["message"] = "\ud800"
        with self.assertRaisesRegex(ValueError, "UTF-8"):
            editing.evaluate_qa_gate(report)

    def test_unpaired_surrogate_in_composite_criteria(self):
        with self.assertRaisesRegex(ValueError, "UTF-8"):
            self.combine([], criteria={"model": "\ud800"})

    def test_unpaired_surrogate_in_expected_selector(self):
        with self.assertRaisesRegex(ValueError, "UTF-8"):
            editing.validate_timeline([], expected=[{"text": "\ud800"}])

    def test_reject_composite_reused_as_structural(self):
        combined = self.combine([{"code": "BLACK_FRAME", "severity": "error"}])
        with self.assertRaisesRegex(ValueError, "raw timeline report"):
            editing.combine_qa_reports(combined, {"audio": external([])}, {"audio": {}})

    def test_visual_frame_pair_must_be_complete(self):
        for issue in ({"start_frame": 0}, {"end_frame": 12}):
            with self.subTest(issue=issue), self.assertRaisesRegex(ValueError, "together"):
                self.combine([{"code": "BLACK_FRAME", "severity": "error", **issue}])

    def test_visual_frame_pair_must_be_strict_integer(self):
        for start, end in ((4, 3), (True, 10), (-1, 5), (0, 2.5),
                           (0, editing.MAX_FRAME + 1)):
            with self.subTest(start=start, end=end), self.assertRaisesRegex(
                    ValueError, "frame interval"):
                self.combine([{"code": "BLACK_FRAME", "severity": "error",
                               "start_frame": start, "end_frame": end}])

    def test_conflicting_visual_frame_range_rejected(self):
        with self.assertRaisesRegex(ValueError, "frame interval"):
            self.combine([{"code": "BLACK_FRAME", "severity": "error",
                           "start_frame": 0, "end_frame": 10, "frame_range": [1, 10]}])

    def test_valid_visual_interval_normalized(self):
        result = self.combine([{"code": "BLACK_FRAME", "severity": "error",
                                "start_frame": 0, "end_frame": 10}])
        self.assertEqual(result["issues"][0]["frame_range"], [0, 10])
        self.assertEqual(editing.evaluate_qa_gate(result)["reason_code"], "QA_ISSUES_REMAIN")

    def test_single_frame_visual_interval_supported(self):
        result = self.combine([{"code": "BLACK_FRAME", "severity": "error",
                                "start_frame": 12, "end_frame": 12}])
        self.assertEqual(result["issues"][0]["frame_range"], [12, 12])

    def test_external_frame_range_rejects_malformed_bounds(self):
        for span in ([3, 2], [True, 4], [3], [1, "5"]):
            with self.subTest(span=span), self.assertRaisesRegex(ValueError, "frame_range"):
                self.combine([{"code": "CLIPPING", "severity": "error",
                               "frame_range": span}], source="audio")

    def test_invalid_external_qa_code_rejected(self):
        for code in (" ", "BAD CODE", "BAD\nCODE", "\ud800", "X" * 129):
            with self.subTest(code=repr(code)), self.assertRaisesRegex(ValueError, "invalid QA issue"):
                self.combine([{"code": code, "severity": "error"}])

    def test_invalid_structural_qa_code_rejected(self):
        report = editing.validate_timeline([], expected=[{"item_id": "missing"}])
        report["issues"][0]["code"] = "INVALID CODE"
        with self.assertRaisesRegex(ValueError, "require a code"):
            editing.evaluate_qa_gate(report)


if __name__ == "__main__":
    unittest.main()
