"""Opt-in sustained coverage and aggregate flash QA regressions (#61)."""
import unittest

import scene_qa


def scene(name, start, end):
    return {"id": name, "start_frame": start, "end_frame": end}


def item(name, start, length, kind="VideoItem"):
    return {"item_id": name, "frame": start, "length": length,
            "type": kind, "layer": 1}


class SustainedCoverageTests(unittest.TestCase):
    def inspect(self, items, ranges=None, **options):
        return scene_qa.inspect(items, {
            "scene_ranges": ranges or [scene("s", 0, 20)],
            "require_audio": False,
            "max_visual_gap_frames": 1000,
            **options,
        })

    def codes(self, result):
        return [row["code"] for row in result["issues"]]

    def test_minimum_run_count_rejects_empty_coverage(self):
        result = self.inspect([], min_visual_covered_run_count=1)
        self.assertIn("SCENE_VISUAL_COVERED_RUN_COUNT_LOW", self.codes(result))
        self.assertEqual(result["overall_coverage"]["visual"]["covered_run_count"], 0)

    def test_minimum_run_count_exact_boundary_passes(self):
        result = self.inspect([item("v", 0, 5)],
                              min_visual_covered_run_count=1)
        self.assertTrue(result["passed"])

    def test_minimum_run_count_rejects_insufficient_runs(self):
        result = self.inspect([item("v", 0, 5)],
                              min_visual_covered_run_count=2)
        self.assertIn("SCENE_VISUAL_COVERED_RUN_COUNT_LOW", self.codes(result))

    def test_minimum_run_count_counts_contiguous_scenes_once(self):
        result = self.inspect(
            [item("v", 8, 4)],
            [scene("a", 0, 10), scene("b", 10, 20)],
            min_visual_covered_run_count=2)
        self.assertIn("SCENE_VISUAL_COVERED_RUN_COUNT_LOW", self.codes(result))

    def test_longest_run_rejects_empty_coverage(self):
        result = self.inspect([], min_visual_longest_covered_run_frames=1)
        self.assertIn("SCENE_VISUAL_LONGEST_COVERAGE_RUN_LOW", self.codes(result))
        self.assertIsNone(result["overall_coverage"]["visual"]["longest_covered_range"])

    def test_longest_run_gate_rejects_fragmented_coverage(self):
        result = self.inspect(
            [item("a", 0, 3), item("b", 6, 3)],
            min_visual_longest_covered_run_frames=4)
        self.assertIn("SCENE_VISUAL_LONGEST_COVERAGE_RUN_LOW", self.codes(result))

    def test_longest_run_gate_joins_adjacent_scenes(self):
        result = self.inspect(
            [item("v", 8, 4)],
            [scene("a", 0, 10), scene("b", 10, 20)],
            min_visual_longest_covered_run_frames=4)
        self.assertTrue(result["passed"])
        self.assertEqual(result["overall_coverage"]["visual"]["longest_covered_range"],
                         {"start_frame": 8, "end_frame": 12,
                          "scene_ids": ["a", "b"]})

    def test_longest_run_does_not_bridge_unselected_frames(self):
        result = self.inspect(
            [item("a", 0, 2), item("b", 10, 2)],
            [scene("a", 0, 5), scene("b", 10, 15)],
            min_visual_longest_covered_run_frames=3)
        self.assertIn("SCENE_VISUAL_LONGEST_COVERAGE_RUN_LOW", self.codes(result))

    def test_short_frame_budget_catches_many_small_flashes(self):
        result = self.inspect(
            [item("a", 0, 2), item("b", 5, 2), item("c", 10, 6)],
            min_visual_covered_run_frames=3,
            max_visual_short_covered_run_count=2,
            max_visual_short_covered_run_total_frames=3)
        self.assertIn("SCENE_VISUAL_SHORT_COVERAGE_FRAMES", self.codes(result))
        self.assertNotIn("SCENE_VISUAL_SHORT_COVERAGE_RUN", self.codes(result))
        self.assertEqual(result["overall_coverage"]["visual"]["short_covered_frames"], 4)

    def test_short_frame_budget_exact_boundary_passes(self):
        result = self.inspect(
            [item("a", 0, 2), item("b", 5, 2)],
            min_visual_covered_run_frames=3,
            max_visual_short_covered_run_count=2,
            max_visual_short_covered_run_total_frames=4)
        self.assertTrue(result["passed"])

    def test_short_frame_budget_requires_minimum_duration(self):
        with self.assertRaisesRegex(ValueError, "requires"):
            self.inspect([], max_visual_short_covered_run_total_frames=0)

    def test_new_options_reject_invalid_numeric_values(self):
        for key in ("min_visual_covered_run_count",
                    "min_visual_longest_covered_run_frames"):
            for bad in (0, -1, True, 1.5, "2", None):
                with self.subTest(key=key, bad=bad), self.assertRaises(ValueError):
                    self.inspect([], **{key: bad})
        for bad in (-1, True, 1.5, "2", None):
            with self.subTest(bad=bad), self.assertRaises(ValueError):
                self.inspect([], min_visual_covered_run_frames=3,
                             max_visual_short_covered_run_total_frames=bad)

    def test_disabled_media_rejects_new_requirements(self):
        for key, value in (("min_audio_covered_run_count", 1),
                           ("min_audio_longest_covered_run_frames", 1),
                           ("max_audio_short_covered_run_total_frames", 0)):
            with self.subTest(key=key), self.assertRaisesRegex(ValueError, "disabled audio"):
                self.inspect([], **{key: value})

    def test_audio_longest_gate_independent_from_visual(self):
        result = scene_qa.inspect(
            [item("v", 0, 20), item("a", 0, 2, "VoiceItem")],
            {"scene_ranges": [scene("s", 0, 20)],
             "max_audio_gap_frames": 1000,
             "min_audio_longest_covered_run_frames": 3})
        self.assertIn("SCENE_AUDIO_LONGEST_COVERAGE_RUN_LOW", self.codes(result))
        self.assertNotIn("SCENE_VISUAL_LONGEST_COVERAGE_RUN_LOW", self.codes(result))

    def test_new_diagnostics_remain_opt_in(self):
        result = self.inspect([item("v", 0, 20)])
        self.assertNotIn("longest_covered_range", result["overall_coverage"]["visual"])
        self.assertNotIn("short_covered_frames", result["overall_coverage"]["visual"])


if __name__ == "__main__":
    unittest.main()
