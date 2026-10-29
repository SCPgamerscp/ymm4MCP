"""Regression tests for stable timeline identity and saturated QA scoring (Issue #61)."""
import unittest

import editing


def item(identity, frame=0, layer=1, length=10):
    return {"item_id": identity, "frame": frame, "length": length, "layer": layer}


def missing_expected(count=6):
    return [{"item_id": f"not-present-{i}"} for i in range(count)]


class StableTimelineIdentityTests(unittest.TestCase):
    def test_duplicate_ids_across_layers_fail_without_overlap(self):
        report = editing.validate_timeline([item("native:same", layer=1),
                                            item("native:same", layer=2)])
        self.assertFalse(report["passed"])
        self.assertEqual([issue["code"] for issue in report["issues"]],
                         ["DUPLICATE_ITEM_ID"])
        self.assertEqual(report["issues"][0]["indices"], [0, 1])
        self.assertEqual(report["issues"][0]["item_ids"], ["native:same"])

    def test_three_reuses_are_not_hidden_by_unique_layers(self):
        report = editing.validate_timeline([item("runtime:one", layer=i)
                                            for i in range(3)])
        self.assertEqual([issue["code"] for issue in report["issues"]],
                         ["DUPLICATE_ITEM_ID", "DUPLICATE_ITEM_ID"])
        self.assertEqual([issue["indices"] for issue in report["issues"]],
                         [[0, 1], [0, 2]])

    def test_invalid_supplied_ids_fail_closed(self):
        for invalid in ("", "   ", "bad\\nname", "bad\\x00name",
                        "bad\\x7fname", 12, False, [], {}):
            with self.subTest(invalid=invalid):
                report = editing.validate_timeline([item(invalid)])
                self.assertFalse(report["passed"])
                self.assertEqual(report["issues"][0]["code"], "INVALID_ITEM_ID")

    def test_missing_and_null_ids_are_supported_for_legacy_snapshots(self):
        anonymous = {"frame": 0, "length": 10, "layer": 1}
        null_id = {"item_id": None, "frame": 0, "length": 10, "layer": 2}
        report = editing.validate_timeline([anonymous, null_id])
        self.assertTrue(report["passed"], report["issues"])

    def test_distinct_valid_ids_remain_clean(self):
        report = editing.validate_timeline([item("native:one", layer=1),
                                            item("runtime:two", layer=2)])
        self.assertTrue(report["passed"], report["issues"])

    def test_duplicate_id_is_not_lost_when_expected_is_satisfied(self):
        report = editing.validate_timeline(
            [item("duplicate", layer=1), item("duplicate", layer=2)],
            expected=[{"frame": 0, "layer": 1}])
        self.assertFalse(report["passed"])
        self.assertEqual(report["issues"][0]["code"], "DUPLICATE_ITEM_ID")


class SaturatedQaRegressionTests(unittest.TestCase):
    def test_more_errors_at_zero_score_stops_and_suggests_rollback(self):
        expected = missing_expected()
        before = editing.validate_timeline([], expected=expected)
        after = editing.validate_timeline(
            [item("left", frame=0), item("right", frame=5)],
            expected=expected)
        self.assertEqual(before["score"], 0)
        self.assertEqual(after["score"], 0)
        self.assertEqual(after["summary"]["errors"], before["summary"]["errors"] + 1)
        result = editing.evaluate_qa_gate(after, [before])
        self.assertEqual(result["decision"], "stop")
        self.assertEqual(result["reason_code"], "QA_REGRESSED")
        self.assertEqual(result["suggested_action"], "consider_checkpoint_rollback")
        self.assertEqual(result["score_delta"], 0)

    def test_more_warnings_at_zero_score_stops(self):
        expected = missing_expected()
        before = editing.validate_timeline([], expected=expected)
        after = editing.validate_timeline(
            [item("left", frame=0), item("right", frame=20)],
            expected=expected)
        self.assertEqual((before["score"], after["score"]), (0, 0))
        self.assertEqual(after["summary"]["warnings"], 1)
        self.assertEqual(editing.evaluate_qa_gate(after, [before])["reason_code"],
                         "QA_REGRESSED")

    def test_fewer_errors_at_zero_score_remains_repairable(self):
        expected = missing_expected()
        before = editing.validate_timeline(
            [item("left", frame=0), item("right", frame=5)],
            expected=expected)
        after = editing.validate_timeline([], expected=expected)
        self.assertEqual((before["score"], after["score"]), (0, 0))
        self.assertEqual(editing.evaluate_qa_gate(after, [before])["reason_code"],
                         "QA_ISSUES_REMAIN")

    def test_unchanged_zero_score_is_stalled_not_regressed(self):
        expected = missing_expected()
        before = editing.validate_timeline([], expected=expected)
        after = editing.validate_timeline([], expected=expected)
        self.assertEqual(editing.evaluate_qa_gate(after, [before])["reason_code"],
                         "QA_STALLED")


if __name__ == "__main__":
    unittest.main()
