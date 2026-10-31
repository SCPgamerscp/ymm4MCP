"""Regression tests for scoped edge gaps and consecutive empty scenes (#61)."""
import unittest

import scene_qa


def scene(name, start, end):
    return {"id": name, "start_frame": start, "end_frame": end}


def item(name, start, length, kind="VideoItem"):
    return {"item_id": name, "frame": start, "length": length,
            "type": kind, "layer": 1}


class EdgeAndEmptyStreakTests(unittest.TestCase):
    def inspect(self, items, ranges=None, **options):
        return scene_qa.inspect(items, {
            "scene_ranges": ranges or [scene("a", 0, 10), scene("b", 10, 20)],
            "require_audio": False, "max_visual_gap_frames": 1000,
            **options,
        })

    @staticmethod
    def codes(result):
        return [issue["code"] for issue in result["issues"]]

    def test_leading_gap_joins_touching_selected_scenes(self):
        result = self.inspect([item("v", 12, 8)],
                              max_visual_leading_uncovered_frames=11)
        self.assertIn("SCENE_VISUAL_LEADING_UNCOVERED", self.codes(result))
        self.assertEqual(result["overall_coverage"]["visual"]["leading_uncovered_frames"], 12)

    def test_leading_gap_exact_budget_passes(self):
        result = self.inspect([item("v", 12, 8)],
                              max_visual_leading_uncovered_frames=12)
        self.assertTrue(result["passed"])

    def test_leading_gap_does_not_cross_unselected_time(self):
        result = self.inspect([item("v", 17, 8)],
                              [scene("a", 0, 10), scene("b", 15, 25)],
                              max_visual_leading_uncovered_frames=10)
        self.assertTrue(result["passed"])
        self.assertEqual(result["overall_coverage"]["visual"]["leading_uncovered_frames"], 10)

    def test_trailing_gap_joins_touching_selected_scenes(self):
        result = self.inspect([item("v", 0, 8)],
                              max_visual_trailing_uncovered_frames=11)
        self.assertIn("SCENE_VISUAL_TRAILING_UNCOVERED", self.codes(result))
        self.assertEqual(result["overall_coverage"]["visual"]["trailing_uncovered_frames"], 12)

    def test_trailing_gap_exact_budget_passes(self):
        result = self.inspect([item("v", 0, 8)],
                              max_visual_trailing_uncovered_frames=12)
        self.assertTrue(result["passed"])

    def test_trailing_gap_does_not_cross_unselected_time(self):
        result = self.inspect([item("v", 0, 8)],
                              [scene("a", 0, 10), scene("b", 15, 25)],
                              max_visual_trailing_uncovered_frames=10)
        self.assertTrue(result["passed"])
        self.assertEqual(result["overall_coverage"]["visual"]["trailing_uncovered_frames"], 10)

    def test_fully_empty_scope_counts_both_edges(self):
        result = self.inspect([], max_visual_leading_uncovered_frames=19,
                              max_visual_trailing_uncovered_frames=19)
        self.assertEqual(result["overall_coverage"]["visual"]["leading_uncovered_frames"], 20)
        self.assertEqual(result["overall_coverage"]["visual"]["trailing_uncovered_frames"], 20)
        self.assertIn("SCENE_VISUAL_LEADING_UNCOVERED", self.codes(result))
        self.assertIn("SCENE_VISUAL_TRAILING_UNCOVERED", self.codes(result))

    def test_consecutive_empty_scene_count_fails(self):
        ranges = [scene("a", 0, 10), scene("b", 10, 20),
                  scene("c", 20, 30), scene("d", 30, 40)]
        result = self.inspect([item("v1", 0, 10), item("v2", 30, 10)], ranges,
                              max_visual_consecutive_empty_scene_count=1)
        self.assertIn("SCENE_VISUAL_CONSECUTIVE_EMPTY_SCENES", self.codes(result))
        row = result["overall_coverage"]["visual"]["longest_empty_scene_streak_by_count"]
        self.assertEqual(row["scene_ids"], ["b", "c"])
        self.assertEqual(row["scene_count"], 2)

    def test_consecutive_empty_scene_count_boundary_passes(self):
        result = self.inspect([item("v", 0, 10)],
                              max_visual_consecutive_empty_scene_count=1)
        self.assertTrue(result["passed"])

    def test_empty_streak_does_not_cross_unselected_time(self):
        ranges = [scene("a", 0, 10), scene("b", 15, 25)]
        result = self.inspect([], ranges, max_visual_consecutive_empty_scene_count=1,
                              max_visual_consecutive_empty_scene_frames=10)
        self.assertTrue(result["passed"])
        row = result["overall_coverage"]["visual"]["longest_empty_scene_streak_by_count"]
        self.assertEqual(row["scene_count"], 1)

    def test_consecutive_empty_frame_budget_detects_long_scene(self):
        ranges = [scene("a", 0, 5), scene("b", 5, 25), scene("c", 25, 30)]
        result = self.inspect([item("v", 25, 5)], ranges,
                              max_visual_consecutive_empty_scene_frames=24)
        self.assertIn("SCENE_VISUAL_CONSECUTIVE_EMPTY_FRAMES", self.codes(result))
        row = result["overall_coverage"]["visual"]["longest_empty_scene_streak_by_frames"]
        self.assertEqual(row["frames"], 25)

    def test_consecutive_empty_frame_budget_exact_boundary_passes(self):
        ranges = [scene("a", 0, 5), scene("b", 5, 25), scene("c", 25, 30)]
        result = self.inspect([item("v", 25, 5)], ranges,
                              max_visual_consecutive_empty_scene_frames=25)
        self.assertTrue(result["passed"])

    def test_empty_streak_count_and_frames_select_different_streaks(self):
        ranges = [scene("a", 0, 4), scene("b", 4, 8),
                  scene("c", 8, 9), scene("d", 9, 30)]
        result = self.inspect([item("v", 8, 1)], ranges,
                              max_visual_consecutive_empty_scene_count=1,
                              max_visual_consecutive_empty_scene_frames=20)
        overall = result["overall_coverage"]["visual"]
        self.assertEqual(overall["longest_empty_scene_streak_by_count"]["scene_ids"], ["a", "b"])
        self.assertEqual(overall["longest_empty_scene_streak_by_frames"]["scene_ids"], ["d"])
        self.assertIn("SCENE_VISUAL_CONSECUTIVE_EMPTY_SCENES", self.codes(result))
        self.assertIn("SCENE_VISUAL_CONSECUTIVE_EMPTY_FRAMES", self.codes(result))

    def test_audio_edge_is_independent_of_visual(self):
        result = scene_qa.inspect(
            [item("v", 0, 20), item("voice", 10, 10, "VoiceItem")],
            {"scene_ranges": [scene("a", 0, 20)],
             "max_audio_gap_frames": 1000,
             "max_audio_leading_uncovered_frames": 9})
        self.assertIn("SCENE_AUDIO_LEADING_UNCOVERED", self.codes(result))
        self.assertNotIn("SCENE_VISUAL_LEADING_UNCOVERED", self.codes(result))

    def test_audio_streak_is_independent_of_visual(self):
        result = scene_qa.inspect(
            [item("v", 0, 20)],
            {"scene_ranges": [scene("a", 0, 10), scene("b", 10, 20)],
             "max_audio_gap_frames": 1000,
             "max_audio_consecutive_empty_scene_count": 1})
        self.assertIn("SCENE_AUDIO_CONSECUTIVE_EMPTY_SCENES", self.codes(result))
        self.assertNotIn("SCENE_VISUAL_CONSECUTIVE_EMPTY_SCENES", self.codes(result))

    def test_new_diagnostics_are_opt_in(self):
        result = self.inspect([item("v", 0, 20)])
        overall = result["overall_coverage"]["visual"]
        self.assertNotIn("leading_uncovered_frames", overall)
        self.assertNotIn("trailing_uncovered_frames", overall)
        self.assertNotIn("longest_empty_scene_streak_by_count", overall)

    def test_new_options_reject_invalid_numbers(self):
        keys = ("max_visual_leading_uncovered_frames",
                "max_visual_trailing_uncovered_frames",
                "max_visual_consecutive_empty_scene_count",
                "max_visual_consecutive_empty_scene_frames")
        for key in keys:
            for invalid in (-1, True, 1.5, "1", None):
                with self.subTest(key=key, invalid=invalid):
                    with self.assertRaises(ValueError):
                        self.inspect([], **{key: invalid})

    def test_disabled_audio_rejects_active_new_constraints(self):
        keys = ("max_audio_leading_uncovered_frames",
                "max_audio_trailing_uncovered_frames",
                "max_audio_consecutive_empty_scene_count",
                "max_audio_consecutive_empty_scene_frames")
        for key in keys:
            with self.subTest(key=key):
                with self.assertRaisesRegex(ValueError, "disabled audio"):
                    self.inspect([], **{key: 0})


if __name__ == "__main__":
    unittest.main()
