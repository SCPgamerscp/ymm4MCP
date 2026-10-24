"""Aggregate scene gap budgets should not be bypassed by fragmented small gaps."""
import unittest

import scene_qa


RANGES = [{"id": "intro", "start_frame": 0, "end_frame": 30}]
ITEMS = [{"item_id": f"v{i}", "type": "VideoItem", "frame": start,
          "length": 4, "layer": 1} for i, start in enumerate((4, 12, 20))]


class SceneTotalCoverageTests(unittest.TestCase):
    def config(self, **options):
        return {"scene_ranges": RANGES, "require_audio": False,
                "max_visual_gap_frames": 6, **options}

    def test_fragmented_gaps_can_exceed_total_budget(self):
        result = scene_qa.inspect(ITEMS, self.config(max_visual_total_uncovered_frames=17))
        self.assertFalse(result["passed"])
        self.assertEqual(len(result["issues"]), 1)
        self.assertEqual(result["issues"][0]["code"], "SCENE_VISUAL_GAP")
        self.assertEqual(result["issues"][0]["uncovered_ranges"], [
            {"start_frame": 0, "end_frame": 4},
            {"start_frame": 8, "end_frame": 12},
            {"start_frame": 16, "end_frame": 20},
            {"start_frame": 24, "end_frame": 30}])

    def test_budget_equality_and_omission_keep_existing_pass(self):
        self.assertTrue(scene_qa.inspect(ITEMS, self.config())["passed"])
        self.assertTrue(scene_qa.inspect(ITEMS, self.config(
            max_visual_total_uncovered_frames=18))["passed"])

    def test_audio_budget_is_independent(self):
        items = ITEMS + [{"item_id": "a", "type": "AudioItem", "frame": 0,
                          "length": 30, "layer": 2}]
        result = scene_qa.inspect(items, {"scene_ranges": RANGES,
                                         "max_visual_gap_frames": 6,
                                         "max_visual_total_uncovered_frames": 17,
                                         "max_audio_total_uncovered_frames": 0})
        self.assertEqual([i["code"] for i in result["issues"]], ["SCENE_VISUAL_GAP"])

    def test_invalid_budgets_are_rejected(self):
        for value in (-1, True, 1.5, "10", None):
            with self.subTest(value=value), self.assertRaises(ValueError):
                scene_qa.inspect(ITEMS, self.config(max_visual_total_uncovered_frames=value))

    def test_single_issue_when_both_limits_fail(self):
        result = scene_qa.inspect(ITEMS, self.config(max_visual_gap_frames=1,
                                                    max_visual_total_uncovered_frames=10))
        self.assertEqual(len(result["issues"]), 1)
        self.assertEqual(len(result["issues"][0]["uncovered_ranges"]), 4)


if __name__ == "__main__":
    unittest.main()
