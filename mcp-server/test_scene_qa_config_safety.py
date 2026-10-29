"""Scene QA configuration must not silently ignore requested coverage checks."""
import unittest

import scene_qa


RANGES = [{"id": "intro", "start_frame": 0, "end_frame": 10}]
ITEMS = [{"item_id": "image", "type": "ImageItem", "frame": 0,
          "length": 5, "layer": 1}]


class SceneQaConfigurationSafetyTests(unittest.TestCase):
    def test_string_ratios_are_rejected_instead_of_coerced(self):
        for value in ("0.5", "1", "0", "", None, [], {}):
            with self.subTest(value=value), self.assertRaises(ValueError):
                scene_qa.inspect(ITEMS, {"scene_ranges": RANGES,
                                         "min_visual_coverage_ratio": value})

    def test_disabled_visual_ratio_cannot_silently_pass(self):
        with self.assertRaisesRegex(ValueError, "disabled visual coverage"):
            scene_qa.inspect(ITEMS, {"scene_ranges": RANGES,
                                     "require_visual": False,
                                     "min_visual_coverage_ratio": 1.0})

    def test_disabled_audio_total_budget_cannot_silently_pass(self):
        with self.assertRaisesRegex(ValueError, "disabled audio coverage"):
            scene_qa.inspect(ITEMS, {"scene_ranges": RANGES,
                                     "require_audio": False,
                                     "max_audio_total_uncovered_frames": 0})

    def test_disabled_category_without_extra_requirements_still_works(self):
        result = scene_qa.inspect(ITEMS, {"scene_ranges": RANGES,
                                          "require_audio": False,
                                          "max_visual_gap_frames": 5})
        self.assertTrue(result["passed"])

    def test_valid_numeric_ratio_is_still_enforced(self):
        result = scene_qa.inspect(ITEMS, {"scene_ranges": RANGES,
                                          "require_audio": False,
                                          "max_visual_gap_frames": 10,
                                          "min_visual_coverage_ratio": 0.6})
        self.assertFalse(result["passed"])
        self.assertIn("SCENE_VISUAL_COVERAGE_LOW",
                      [issue["code"] for issue in result["issues"]])


if __name__ == "__main__":
    unittest.main()
