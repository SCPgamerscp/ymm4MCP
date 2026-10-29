"""Regression coverage for composite QA omissions and subtitle/repair safety (#61)."""
import copy
import unittest

import editing


def external(issues):
    return {"success": True, "passed": not any(i["severity"] == "error" for i in issues),
            "issues": issues}


def overflowing_structural(count):
    items = [{"frame": 0, "length": 2, "layer": i, "item_id": f"id-{i}"}
             for i in range(count)]
    return editing.validate_timeline(items, duration=1)


class CompositeOmissionAndSubtitleTests(unittest.TestCase):
    def test_composite_preserves_omitted_structural_issues(self):
        structural = overflowing_structural(1100)
        prior_omitted = structural["issues"][-1]["omitted_count"]
        self.assertEqual(prior_omitted, 101)
        combined = editing.combine_qa_reports(
            structural, {"visual": external([{"code": "BLACK_FRAME", "severity": "error"}])},
            {"visual": {"frames": 1}})
        self.assertEqual(len(combined["issues"]), 1000)
        self.assertEqual(combined["issues"][-1]["omitted_count"], prior_omitted + 1)
        self.assertEqual(editing.evaluate_qa_gate(combined)["reason_code"], "QA_ISSUES_REMAIN")

    def test_composite_retains_dropped_external_marker_omissions(self):
        visual = [{"code": "BLACK_FRAME", "severity": "error"} for _ in range(999)]
        visual.append({"code": "QA_ISSUES_TRUNCATED", "severity": "error",
                       "omitted_count": 50})
        result = editing.combine_qa_reports(
            editing.validate_timeline([]),
            {"visual": external(visual),
             "audio": external([{"code": "CLIPPING", "severity": "error"}])},
            {"visual": {}, "audio": {}})
        self.assertEqual(result["issues"][-1]["omitted_count"], 51)

    def test_nested_omissions_worsening_triggers_rollback_guidance(self):
        def composite(count):
            return editing.combine_qa_reports(
                overflowing_structural(count),
                {"visual": external([{"code": "BLACK_FRAME", "severity": "error"}])},
                {"visual": {"frames": 1}})
        before = composite(1100)
        after = composite(1200)
        self.assertEqual(before["score"], after["score"])
        self.assertEqual(editing.evaluate_qa_gate(after, [before])["reason_code"],
                         "QA_REGRESSED")

    def test_changed_item_count_prevents_false_stall(self):
        expected = [{"item_id": "not-here"}]
        before = editing.validate_timeline(
            [{"frame": 0, "length": 1, "layer": 0}], expected=expected)
        after = editing.validate_timeline(
            [{"frame": 0, "length": 1, "layer": 0},
             {"frame": 0, "length": 1, "layer": 1}], expected=expected)
        self.assertEqual(before["issues"], after["issues"])
        self.assertEqual(editing.evaluate_qa_gate(after, [before])["reason_code"],
                         "QA_ISSUES_REMAIN")

    def test_identical_item_count_and_issues_still_stalls(self):
        report = editing.validate_timeline([], expected=[{"item_id": "missing"}])
        self.assertEqual(editing.evaluate_qa_gate(report, [copy.deepcopy(report)])["reason_code"],
                         "QA_STALLED")

    def test_reject_invalid_item_count_evidence(self):
        report = editing.validate_timeline([], expected=[{"item_id": "missing"}])
        for invalid in (-1, True, 1.5, "3"):
            altered = copy.deepcopy(report)
            altered["item_count"] = invalid
            with self.subTest(invalid=invalid), self.assertRaisesRegex(ValueError, "item_count"):
                editing.evaluate_qa_gate(altered)

    @staticmethod
    def timeline(extra_text, extra_frame=0):
        return [
            {"frame": 0, "length": 10, "layer": 0, "type": "VoiceItem",
             "text": "Hello", "item_id": "voice"},
            {"frame": 0, "length": 10, "layer": 1, "type": "TextItem",
             "text": "Hello", "item_id": "correct"},
            {"frame": extra_frame, "length": 10, "layer": 2, "type": "TextItem",
             "text": extra_text, "item_id": "other"},
        ]

    def test_conflicting_subtitle_not_masked_by_correct_subtitle(self):
        report = editing.validate_timeline(
            self.timeline("Wrong"), subtitle_layers=[1, 2], include_gaps=False)
        mismatches = [i for i in report["issues"] if i["code"] == "SUBTITLE_TEXT_MISMATCH"]
        self.assertEqual(len(mismatches), 1)
        self.assertEqual(mismatches[0]["actual_texts"], ["Wrong"])
        self.assertFalse(report["passed"])

    def test_multiple_identical_subtitles_remain_valid(self):
        report = editing.validate_timeline(
            self.timeline("Hello"), subtitle_layers=[1, 2], include_gaps=False)
        self.assertTrue(report["passed"])

    def test_nonoverlapping_unrelated_subtitle_does_not_fail(self):
        report = editing.validate_timeline(
            self.timeline("Wrong", extra_frame=10), subtitle_layers=[1, 2],
            include_gaps=False)
        self.assertTrue(report["passed"])


if __name__ == "__main__":
    unittest.main()
