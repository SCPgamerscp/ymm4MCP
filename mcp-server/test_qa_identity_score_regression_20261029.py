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
        for invalid in ("", "   ", "bad\nname", "bad\x00name",
                        "bad\x7fname", 12, False, [], {}):
            with self.subTest(invalid=invalid):
                report = editing.validate_timeline([item(invalid)])
                self.assertFalse(report["passed"])
                self.assertEqual(report["issues"][0]["code"], "INVALID_ITEM_ID")

    def test_invalid_id_does_not_poison_overlap_repair_report(self):
        report = editing.validate_timeline([
            item("bad\nidentity", frame=0), item("valid", frame=5)])
        self.assertEqual([issue["code"] for issue in report["issues"]],
                         ["INVALID_ITEM_ID", "OVERLAP"])
        self.assertEqual(report["issues"][1]["item_ids"], ["valid"])
        self.assertEqual(editing.evaluate_qa_gate(report)["decision"], "repair")

    def test_invalid_timing_does_not_hide_duplicate_identity(self):
        # A corrupt host item can still reuse an ID belonging to a valid item.
        # The collision must be diagnosed regardless of which appears first.
        good = item("native:shared", layer=1)
        broken = {**item("native:shared", layer=2), "frame": "invalid"}
        for items in ([good, broken], [broken, good]):
            with self.subTest(first_is_valid=items[0] is good):
                report = editing.validate_timeline(items)
                self.assertFalse(report["passed"])
                self.assertEqual(
                    sorted(issue["code"] for issue in report["issues"]),
                    ["DUPLICATE_ITEM_ID", "INVALID_ITEM"])
                collision = next(issue for issue in report["issues"]
                                 if issue["code"] == "DUPLICATE_ITEM_ID")
                self.assertEqual(collision["indices"], [0, 1])
                self.assertEqual(collision["item_ids"], ["native:shared"])

    def test_malformed_timing_and_id_are_both_reported(self):
        report = editing.validate_timeline([
            {**item("bad\nidentifier"), "length": 0}])
        self.assertEqual(
            sorted(issue["code"] for issue in report["issues"]),
            ["INVALID_ITEM", "INVALID_ITEM_ID"])

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

    def test_reducing_anonymous_selector_matches_is_progress(self):
        # The item count changes, but both reports still contain the same
        # EXPECTED_AMBIGUOUS issue and have the same score. Match positions
        # alone must not determine whether a repair was productive.
        anonymous = lambda layer: {"frame": 0, "length": 10, "layer": layer}
        before = editing.validate_timeline(
            [anonymous(1), anonymous(2), anonymous(3)],
            expected=[{"length": 10}])
        after = editing.validate_timeline(
            [anonymous(1), anonymous(2)],
            expected=[{"length": 10}])
        self.assertEqual([i["code"] for i in before["issues"]], ["EXPECTED_AMBIGUOUS"])
        self.assertEqual([i["code"] for i in after["issues"]], ["EXPECTED_AMBIGUOUS"])
        self.assertEqual((len(before["issues"][0]["matches"]),
                          len(after["issues"][0]["matches"])), (3, 2))
        self.assertEqual(before["score"], after["score"])
        result = editing.evaluate_qa_gate(after, [before])
        self.assertEqual(result["decision"], "repair")
        self.assertEqual(result["reason_code"], "QA_ISSUES_REMAIN")

    def test_reordering_anonymous_matches_does_not_reset_stall(self):
        anonymous = lambda layer: {"frame": 0, "length": 10, "layer": layer}
        unrelated = {"frame": 0, "length": 5, "layer": 4}
        before = editing.validate_timeline(
            [anonymous(1), anonymous(2), unrelated],
            expected=[{"length": 10}])
        after = editing.validate_timeline(
            [unrelated, anonymous(1), anonymous(2)],
            expected=[{"length": 10}])
        self.assertEqual(before["issues"][0]["matches"], [0, 1])
        self.assertEqual(after["issues"][0]["matches"], [1, 2])
        self.assertEqual(editing.evaluate_qa_gate(after, [before])["reason_code"],
                         "QA_STALLED")

    def test_unrelated_item_count_change_does_not_reset_stall(self):
        # Repair loops must not evade QA_STALLED by adding/removing unrelated
        # timeline items while leaving the actual overlap untouched.
        broken = [item("left", frame=0), item("right", frame=5)]
        unrelated = item("other", frame=50, layer=8)
        before = editing.validate_timeline(broken, include_gaps=False)
        after = editing.validate_timeline([*broken, unrelated], include_gaps=False)
        self.assertEqual(before["item_count"], 2)
        self.assertEqual(after["item_count"], 3)
        self.assertEqual(before["issues"], after["issues"])
        self.assertEqual(
            editing.evaluate_qa_gate(after, [before])["reason_code"],
            "QA_STALLED")
        reverse = editing.evaluate_qa_gate(before, [after])
        self.assertEqual(reverse["reason_code"], "QA_STALLED")

    def test_unchanged_zero_score_is_stalled_not_regressed(self):
        expected = missing_expected()
        before = editing.validate_timeline([], expected=expected)
        after = editing.validate_timeline([], expected=expected)
        self.assertEqual(editing.evaluate_qa_gate(after, [before])["reason_code"],
                         "QA_STALLED")


if __name__ == "__main__":
    unittest.main()
