"""Regression tests for opt-in short media flashes and fragmented occupancy (#61)."""
import unittest

import scene_qa


def scene(name, start, end):
    return {"id": name, "start_frame": start, "end_frame": end}


def item(name, start, length, kind="VideoItem"):
    return {"item_id": name, "frame": start, "length": length,
            "type": kind, "layer": 1}


class CoveredRunTests(unittest.TestCase):
    def inspect(self, items, ranges, **options):
        return scene_qa.inspect(items, {
            "scene_ranges": ranges, "require_audio": False,
            "max_visual_gap_frames": 1000, **options,
        })

    def test_short_flash_is_detected_without_a_gap_failure(self):
        result = self.inspect([item("flash", 5, 2)], [scene("s", 0, 10)],
                              min_visual_covered_run_frames=3)
        self.assertEqual([row["code"] for row in result["issues"]],
                         ["SCENE_VISUAL_SHORT_COVERAGE_RUN"])
        issue = result["issues"][0]
        self.assertEqual(issue["frame_range"], [5, 7])
        self.assertEqual(issue["covered_frames"], 2)
        self.assertEqual(issue["short_coverage_run_count"], 1)

    def test_exact_minimum_run_is_accepted(self):
        result = self.inspect([item("flash", 5, 3)], [scene("s", 0, 10)],
                              min_visual_covered_run_frames=3)
        self.assertTrue(result["passed"])

    def test_contiguous_scenes_join_one_covered_run(self):
        result = self.inspect([item("v", 0, 20)],
                              [scene("a", 0, 10), scene("b", 10, 20)],
                              min_visual_covered_run_frames=15,
                              max_visual_covered_run_count=1)
        self.assertTrue(result["passed"])

    def test_unselected_time_never_bridges_coverage(self):
        result = self.inspect([item("a", 0, 5), item("b", 10, 5)],
                              [scene("a", 0, 5), scene("b", 10, 15)],
                              min_visual_covered_run_frames=6,
                              max_visual_covered_run_count=1)
        self.assertEqual({row["code"] for row in result["issues"]},
                         {"SCENE_VISUAL_SHORT_COVERAGE_RUN",
                          "SCENE_VISUAL_COVERED_RUN_COUNT"})
        issue = next(row for row in result["issues"]
                     if row["code"] == "SCENE_VISUAL_SHORT_COVERAGE_RUN")
        self.assertEqual(issue["short_coverage_run_count"], 2)
        self.assertEqual(issue["scene_ids"], ["a"])

    def test_gaps_split_coverage_even_inside_one_scene(self):
        result = self.inspect([item("a", 0, 3), item("b", 5, 3)],
                              [scene("s", 0, 10)],
                              max_visual_covered_run_count=1)
        self.assertEqual([row["code"] for row in result["issues"]],
                         ["SCENE_VISUAL_COVERED_RUN_COUNT"])
        self.assertEqual(result["issues"][0]["covered_run_count"], 2)

    def test_overlapping_items_do_not_inflate_run_count(self):
        result = self.inspect([item("a", 0, 6), item("b", 3, 7)],
                              [scene("s", 0, 10)],
                              max_visual_covered_run_count=1)
        self.assertTrue(result["passed"])

    def test_empty_media_has_no_short_covered_run(self):
        result = self.inspect([], [scene("s", 0, 10)],
                              min_visual_covered_run_frames=5,
                              max_visual_covered_run_count=0)
        self.assertTrue(result["passed"])

    def test_audio_coverage_is_independent_of_visual(self):
        result = scene_qa.inspect(
            [item("v", 0, 10), item("voice", 4, 2, "VoiceItem")],
            {"scene_ranges": [scene("s", 0, 10)],
             "require_visual": False, "max_audio_gap_frames": 1000,
             "min_audio_covered_run_frames": 3})
        self.assertEqual([row["code"] for row in result["issues"]],
                         ["SCENE_AUDIO_SHORT_COVERAGE_RUN"])

    def test_invalid_minimums_and_counts_fail_closed(self):
        for key in ("min_visual_covered_run_frames",
                    "max_visual_covered_run_count"):
            for value in (True, -1, 1.5, "2", None):
                with self.subTest(key=key, value=value), self.assertRaises(ValueError):
                    self.inspect([], [scene("s", 0, 10)], **{key: value})
        with self.assertRaises(ValueError):
            self.inspect([], [scene("s", 0, 10)],
                         min_visual_covered_run_frames=0)

    def test_disabled_category_rejects_active_run_requirements(self):
        for key in ("min_audio_covered_run_frames",
                    "max_audio_covered_run_count"):
            with self.subTest(key=key), self.assertRaisesRegex(
                    ValueError, "disabled audio coverage"):
                self.inspect([], [scene("s", 0, 10)], **{key: 1})


if __name__ == "__main__":
    unittest.main()
