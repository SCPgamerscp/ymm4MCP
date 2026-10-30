"""Regression coverage for fragmented and widespread scene media outages (#61)."""
import unittest

import scene_qa


def scene(name, start, end):
    return {"id": name, "start_frame": start, "end_frame": end}


def item(name, start, length, kind="VideoItem"):
    return {"item_id": name, "frame": start, "length": length,
            "type": kind, "layer": 1}


class SceneGapFragmentationTests(unittest.TestCase):
    def inspect(self, items, ranges, **options):
        return scene_qa.inspect(items, {
            "scene_ranges": ranges,
            "require_audio": False,
            "max_visual_gap_frames": 1000,
            **options,
        })

    def test_scene_fragmentation_detected_when_each_gap_is_short(self):
        result = self.inspect(
            [item("a", 0, 3), item("b", 5, 3)],
            [scene("intro", 0, 10)], max_visual_gap_count=1)
        self.assertEqual([row["code"] for row in result["issues"]],
                         ["SCENE_VISUAL_GAP_COUNT"])
        self.assertEqual(result["issues"][0]["gap_count"], 2)
        self.assertEqual(result["issues"][0]["scene_id"], "intro")

    def test_scene_gap_count_exact_limit_passes(self):
        result = self.inspect(
            [item("a", 0, 3), item("b", 5, 3)],
            [scene("intro", 0, 10)], max_visual_gap_count=2)
        self.assertTrue(result["passed"])

    def test_fully_covered_scene_has_zero_gaps(self):
        result = self.inspect([item("a", 0, 10)],
                              [scene("intro", 0, 10)], max_visual_gap_count=0)
        self.assertTrue(result["passed"])
        self.assertEqual(result["overall_coverage"]["visual"]["continuous_gap_count"], 0)

    def test_global_gap_count_accumulates_independent_outages(self):
        result = self.inspect(
            [item("a", 0, 3), item("b", 5, 3), item("c", 10, 3)],
            [scene("intro", 0, 10), scene("outro", 10, 20)],
            max_visual_overall_gap_count=2)
        self.assertEqual([row["code"] for row in result["issues"]],
                         ["SCENE_VISUAL_OVERALL_GAP_COUNT"])
        self.assertEqual(result["issues"][0]["gap_count"], 3)

    def test_boundary_touching_gaps_count_as_one_outage(self):
        result = self.inspect(
            [item("a", 0, 4), item("b", 16, 4)],
            [scene("a", 0, 10), scene("b", 10, 20)],
            max_visual_overall_gap_count=1)
        self.assertTrue(result["passed"])
        self.assertEqual(result["overall_coverage"]["visual"]["continuous_gap_count"], 1)
        self.assertEqual(result["overall_coverage"]["visual"]["uncovered_scene_count"], 2)
        self.assertEqual(result["overall_coverage"]["visual"]["longest_uncovered_range"],
                         {"start_frame": 4, "end_frame": 16, "scene_ids": ["a", "b"]})

    def test_unselected_time_breaks_continuous_gap_count(self):
        result = self.inspect(
            [], [scene("a", 0, 5), scene("b", 10, 15)],
            max_visual_overall_gap_count=1)
        self.assertFalse(result["passed"])
        self.assertEqual(result["overall_coverage"]["visual"]["continuous_gap_count"], 2)

    def test_uncovered_scene_budget_detects_widespread_outages(self):
        result = self.inspect(
            [item("a", 0, 9), item("b", 10, 9), item("c", 20, 10)],
            [scene("a", 0, 10), scene("b", 10, 20), scene("c", 20, 30)],
            max_visual_uncovered_scene_count=1)
        self.assertEqual([row["code"] for row in result["issues"]],
                         ["SCENE_VISUAL_UNCOVERED_SCENE_COUNT"])
        self.assertEqual(result["issues"][0]["uncovered_scene_count"], 2)

    def test_uncovered_scene_budget_exact_limit_passes(self):
        result = self.inspect(
            [item("a", 0, 9), item("b", 10, 10)],
            [scene("a", 0, 10), scene("b", 10, 20)],
            max_visual_uncovered_scene_count=1)
        self.assertTrue(result["passed"])

    def test_audio_gap_budgets_are_independent(self):
        result = scene_qa.inspect(
            [item("video", 0, 20), item("voice", 0, 4, "VoiceItem"),
             item("voice2", 6, 4, "VoiceItem")],
            {"scene_ranges": [scene("a", 0, 10), scene("b", 10, 20)],
             "require_visual": False, "max_audio_gap_frames": 1000,
             "max_audio_overall_gap_count": 1})
        self.assertEqual([row["code"] for row in result["issues"]],
                         ["SCENE_AUDIO_OVERALL_GAP_COUNT"])

    def test_reject_invalid_count_budgets(self):
        for option in ("max_visual_gap_count", "max_visual_overall_gap_count",
                       "max_visual_uncovered_scene_count"):
            for value in (-1, True, 1.5, "2", None):
                with self.subTest(option=option, value=value):
                    with self.assertRaises(ValueError):
                        self.inspect([], [scene("a", 0, 10)], **{option: value})

    def test_disabled_media_cannot_hide_new_constraints(self):
        for option in ("max_audio_gap_count", "max_audio_overall_gap_count",
                       "max_audio_uncovered_scene_count"):
            with self.subTest(option=option):
                with self.assertRaisesRegex(ValueError, "disabled audio coverage"):
                    self.inspect([], [scene("a", 0, 10)], **{option: 1})

    def test_default_gap_statistics_do_not_add_errors(self):
        result = self.inspect(
            [item("a", 0, 4), item("b", 16, 4)],
            [scene("a", 0, 10), scene("b", 10, 20)])
        self.assertTrue(result["passed"])
        self.assertEqual(result["issues"], [])
        self.assertEqual(result["overall_coverage"]["visual"]["continuous_gap_count"], 1)
        self.assertEqual(result["overall_coverage"]["visual"]["uncovered_scene_count"], 2)


if __name__ == "__main__":
    unittest.main()
