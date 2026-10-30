"""Cross-scene and weighted scene QA regression tests (#61)."""
import unittest

import scene_qa


def scene(scene_id, start, end):
    return {"id": scene_id, "start_frame": start, "end_frame": end}


def media(identity, start, duration, kind):
    return {"item_id": identity, "frame": start, "length": duration,
            "type": kind, "layer": 1}


class OverallSceneCoverageTests(unittest.TestCase):
    def inspect(self, items, ranges, **options):
        return scene_qa.inspect(items, {"scene_ranges": ranges,
                                        "max_visual_gap_frames": 1000,
                                        "max_audio_gap_frames": 1000, **options})

    def test_continuous_gap_crossing_scene_boundary_is_detected(self):
        result = self.inspect(
            [media("v", 0, 5, "VideoItem")],
            [scene("a", 0, 10), scene("b", 10, 20)],
            require_audio=False, max_visual_continuous_gap_frames=10)
        issue = next(i for i in result["issues"]
                     if i["code"] == "SCENE_VISUAL_CONTINUOUS_GAP")
        self.assertEqual(issue["frame_range"], [5, 20])
        self.assertEqual(issue["scene_ids"], ["a", "b"])
        self.assertEqual(issue["gap_frames"], 15)

    def test_continuous_gap_exact_limit_passes(self):
        result = self.inspect(
            [media("v", 0, 5, "VideoItem")],
            [scene("a", 0, 10), scene("b", 10, 20)],
            require_audio=False, max_visual_continuous_gap_frames=15)
        self.assertTrue(result["passed"])

    def test_unselected_time_between_scenes_is_not_a_continuous_gap(self):
        result = self.inspect(
            [], [scene("a", 0, 5), scene("b", 10, 15)],
            require_audio=False, max_visual_continuous_gap_frames=5)
        self.assertTrue(result["passed"])
        self.assertEqual(result["overall_coverage"]["visual"]["longest_uncovered_range"],
                         {"start_frame": 0, "end_frame": 5, "scene_ids": ["a"]})

    def test_overall_uncovered_budget_accumulates_small_scene_gaps(self):
        result = self.inspect(
            [media("a", 0, 7, "VideoItem"),
             media("b", 10, 7, "VideoItem")],
            [scene("a", 0, 10), scene("b", 10, 20)],
            require_audio=False, max_visual_overall_uncovered_frames=5)
        self.assertEqual([i["code"] for i in result["issues"]],
                         ["SCENE_VISUAL_OVERALL_UNCOVERED"])
        self.assertEqual(result["issues"][0]["uncovered_frames"], 6)

    def test_overall_uncovered_budget_exact_limit_passes(self):
        result = self.inspect(
            [media("v", 0, 14, "VideoItem")],
            [scene("a", 0, 10), scene("b", 10, 20)],
            require_audio=False, max_visual_overall_uncovered_frames=6)
        self.assertTrue(result["passed"])

    def test_weighted_ratio_uses_duration_not_scene_count(self):
        result = self.inspect(
            [media("v", 0, 9, "VideoItem")],
            [scene("short", 0, 10), scene("long", 10, 100)],
            require_audio=False, min_visual_overall_coverage_ratio=0.5)
        issue = next(i for i in result["issues"]
                     if i["code"] == "SCENE_VISUAL_OVERALL_COVERAGE_LOW")
        self.assertEqual(issue["covered_frames"], 9)
        self.assertEqual(issue["scoped_frames"], 100)
        self.assertEqual(issue["coverage_ratio"], 0.09)

    def test_overall_ratio_exact_boundary_passes(self):
        result = self.inspect(
            [media("v", 0, 10, "VideoItem")],
            [scene("a", 0, 10), scene("b", 10, 20)],
            require_audio=False, min_visual_overall_coverage_ratio=0.5)
        self.assertTrue(result["passed"])

    def test_unselected_time_does_not_reduce_weighted_coverage(self):
        result = self.inspect(
            [media("v1", 0, 5, "VideoItem"),
             media("v2", 100, 5, "VideoItem")],
            [scene("a", 0, 5), scene("b", 100, 105)],
            require_audio=False, min_visual_overall_coverage_ratio=1.0)
        self.assertTrue(result["passed"])
        self.assertEqual(result["overall_coverage"]["visual"]["scoped_frames"], 10)

    def test_audio_overall_thresholds_are_independent(self):
        result = self.inspect(
            [media("v", 0, 20, "VideoItem"),
             media("voice", 0, 5, "VoiceItem")],
            [scene("a", 0, 10), scene("b", 10, 20)],
            max_audio_overall_uncovered_frames=10,
            min_audio_overall_coverage_ratio=0.5)
        self.assertEqual({i["code"] for i in result["issues"]},
                         {"SCENE_AUDIO_OVERALL_UNCOVERED",
                          "SCENE_AUDIO_OVERALL_COVERAGE_LOW"})
        self.assertEqual(result["overall_coverage"]["visual"]["coverage_ratio"], 1.0)
        self.assertEqual(result["overall_coverage"]["audio"]["coverage_ratio"], 0.25)

    def test_rejects_malformed_overall_ratio_options(self):
        for bad in (True, None, "0.5", -0.1, 1.1, float("nan"), float("inf")):
            with self.subTest(bad=bad), self.assertRaises(ValueError):
                self.inspect([], [scene("a", 0, 10)],
                             min_visual_overall_coverage_ratio=bad)

    def test_rejects_invalid_global_limits(self):
        for option in ("max_audio_continuous_gap_frames",
                       "max_visual_overall_uncovered_frames"):
            for bad in (-1, True, "3", 1.2):
                with self.subTest(option=option, bad=bad), self.assertRaises(ValueError):
                    self.inspect([], [scene("a", 0, 10)], **{option: bad})

    def test_disabled_category_cannot_silently_ignore_global_checks(self):
        for option in ("max_visual_continuous_gap_frames",
                       "max_visual_overall_uncovered_frames",
                       "min_visual_overall_coverage_ratio"):
            with self.subTest(option=option), self.assertRaisesRegex(
                    ValueError, "disabled visual coverage"):
                self.inspect([], [scene("a", 0, 10)],
                             require_visual=False, **{option: 0})

    def test_empty_category_has_full_length_gap(self):
        result = self.inspect([], [scene("a", 0, 10)],
                              require_visual=False, max_audio_continuous_gap_frames=9)
        self.assertEqual(result["issues"][-1]["code"],
                         "SCENE_AUDIO_CONTINUOUS_GAP")
        self.assertEqual(result["issues"][-1]["frame_range"], [0, 10])

    def test_global_metrics_reported_even_without_global_limits(self):
        result = self.inspect([media("v", 0, 5, "VideoItem")],
                              [scene("a", 0, 10)], require_audio=False)
        self.assertEqual(result["overall_coverage"]["visual"],
                         {"scoped_frames": 10, "covered_frames": 5,
                          "uncovered_frames": 5, "coverage_ratio": 0.5,
                          "longest_uncovered_range":
                              {"start_frame": 5, "end_frame": 10,
                               "scene_ids": ["a"]}})


if __name__ == "__main__":
    unittest.main()
