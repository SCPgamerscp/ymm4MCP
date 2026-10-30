"""Regression tests for disabled category gates and tolerated short media runs (#61)."""
import unittest

import scene_qa


def scene(name, start, end):
    return {"id": name, "start_frame": start, "end_frame": end}


def item(name, start, length, kind="VideoItem"):
    return {"item_id": name, "frame": start, "length": length,
            "type": kind, "layer": 1}


class SceneShortRunBudgetTests(unittest.TestCase):
    def inspect(self, items, ranges, **options):
        return scene_qa.inspect(items, {
            "scene_ranges": ranges, "require_audio": False,
            "max_visual_gap_frames": 1000, **options,
        })

    def test_explicit_gap_limit_rejected_when_visual_disabled(self):
        with self.assertRaisesRegex(ValueError, "disabled visual coverage"):
            scene_qa.inspect([], {
                "scene_ranges": [scene("s", 0, 10)],
                "require_visual": False,
                "max_visual_gap_frames": 10,
            })

    def test_explicit_gap_limit_rejected_when_audio_disabled(self):
        with self.assertRaisesRegex(ValueError, "disabled audio coverage"):
            self.inspect([], [scene("s", 0, 10)],
                         max_audio_gap_frames=1)

    def test_default_gap_limit_does_not_disable_a_category(self):
        result = scene_qa.inspect(
            [item("v", 0, 10)],
            {"scene_ranges": [scene("s", 0, 10)],
             "require_audio": False})
        self.assertTrue(result["passed"])

    def test_short_run_budget_exact_limit_passes(self):
        result = self.inspect(
            [item("short", 0, 2), item("long", 5, 10)],
            [scene("s", 0, 15)],
            min_visual_covered_run_frames=3,
            max_visual_short_covered_run_count=1)
        self.assertTrue(result["passed"])

    def test_short_run_budget_exceeded_reports_count(self):
        result = self.inspect(
            [item("a", 0, 2), item("b", 5, 2), item("c", 10, 5)],
            [scene("s", 0, 15)],
            min_visual_covered_run_frames=3,
            max_visual_short_covered_run_count=1)
        issues = [row for row in result["issues"]
                  if row["code"] == "SCENE_VISUAL_SHORT_COVERAGE_RUN"]
        self.assertEqual(len(issues), 1)
        self.assertEqual(issues[0]["short_coverage_run_count"], 2)
        self.assertEqual(issues[0]["maximum_short_coverage_run_count"], 1)

    def test_short_run_budget_zero_preserves_original_behavior(self):
        result = self.inspect(
            [item("short", 0, 2)],
            [scene("s", 0, 10)],
            min_visual_covered_run_frames=3)
        self.assertEqual(
            [row["code"] for row in result["issues"]],
            ["SCENE_VISUAL_SHORT_COVERAGE_RUN"])

    def test_cross_scene_run_is_not_double_counted(self):
        result = self.inspect(
            [item("cross", 8, 4), item("short", 15, 1)],
            [scene("a", 0, 10), scene("b", 10, 20)],
            min_visual_covered_run_frames=3,
            max_visual_short_covered_run_count=1)
        self.assertTrue(result["passed"])

    def test_short_run_budget_requires_minimum_duration(self):
        with self.assertRaisesRegex(ValueError, "requires"):
            self.inspect([], [scene("s", 0, 10)],
                         max_visual_short_covered_run_count=1)

    def test_short_run_budget_rejects_invalid_values(self):
        for bad in (-1, True, 1.5, "1", None):
            with self.subTest(bad=bad), self.assertRaises(ValueError):
                self.inspect([], [scene("s", 0, 10)],
                             min_visual_covered_run_frames=3,
                             max_visual_short_covered_run_count=bad)

    def test_disabled_media_rejects_short_run_budget(self):
        with self.assertRaisesRegex(ValueError, "disabled audio coverage"):
            self.inspect([], [scene("s", 0, 10)],
                         min_audio_covered_run_frames=3,
                         max_audio_short_covered_run_count=1)

    def test_audio_short_run_budget_is_independent(self):
        result = scene_qa.inspect(
            [item("video", 0, 20),
             item("voice_a", 0, 2, "VoiceItem"),
             item("voice_b", 5, 2, "VoiceItem")],
            {"scene_ranges": [scene("s", 0, 20)],
             "max_visual_gap_frames": 1000,
             "max_audio_gap_frames": 1000,
             "min_audio_covered_run_frames": 3,
             "max_audio_short_covered_run_count": 1})
        self.assertEqual(
            [row["code"] for row in result["issues"]],
            ["SCENE_AUDIO_SHORT_COVERAGE_RUN"])


if __name__ == "__main__":
    unittest.main()
