"""Scene coverage ratio thresholds are strict, optional, and independent by media type."""
import unittest

import scene_qa


RANGES = [{"id": "intro", "start_frame": 0, "end_frame": 100}]
ITEMS = [
    {"item_id": "image", "type": "ImageItem", "frame": 0, "length": 80, "layer": 1},
    {"item_id": "voice", "type": "VoiceItem", "frame": 0, "length": 20, "layer": 2},
]


class SceneCoverageRatioTests(unittest.TestCase):
    def test_reports_coverage_ratios_for_both_media_types(self):
        result = scene_qa.inspect(ITEMS, {"scene_ranges": RANGES,
                                          "max_visual_gap_frames": 100,
                                          "max_audio_gap_frames": 100})
        self.assertTrue(result["passed"])
        scene = result["scenes"][0]
        self.assertEqual(scene["visual_coverage_ratio"], 0.8)
        self.assertEqual(scene["audio_coverage_ratio"], 0.2)

    def test_visual_threshold_fails_independently(self):
        result = scene_qa.inspect(ITEMS, {"scene_ranges": RANGES,
                                          "max_visual_gap_frames": 100,
                                          "max_audio_gap_frames": 100,
                                          "min_visual_coverage_ratio": 0.81})
        self.assertFalse(result["passed"])
        self.assertEqual([issue["code"] for issue in result["issues"]],
                         ["SCENE_VISUAL_COVERAGE_LOW"])
        self.assertEqual(result["issues"][0]["covered_frames"], 80)

    def test_exact_boundary_passes(self):
        result = scene_qa.inspect(ITEMS, {"scene_ranges": RANGES,
                                          "max_visual_gap_frames": 100,
                                          "max_audio_gap_frames": 100,
                                          "min_visual_coverage_ratio": 0.8,
                                          "min_audio_coverage_ratio": 0.2})
        self.assertTrue(result["passed"])

    def test_audio_threshold_fails_without_visual_failure(self):
        result = scene_qa.inspect(ITEMS, {"scene_ranges": RANGES,
                                          "require_visual": False,
                                          "max_audio_gap_frames": 100,
                                          "min_audio_coverage_ratio": 0.25})
        self.assertEqual([issue["code"] for issue in result["issues"]],
                         ["SCENE_AUDIO_COVERAGE_LOW"])

    def test_invalid_ratios_are_rejected(self):
        for value in (-0.1, 1.1, True, float("nan"), float("inf"), None):
            with self.subTest(value=value), self.assertRaises(ValueError):
                scene_qa.inspect(ITEMS, {"scene_ranges": RANGES,
                                         "min_visual_coverage_ratio": value})


if __name__ == "__main__":
    unittest.main()
