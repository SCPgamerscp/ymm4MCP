"""Regression tests for very small positive script planning speeds."""
import unittest

from editing import plan_script


class ScriptPlanTinySpeedTests(unittest.TestCase):
    def test_subnormal_speed_raises_value_error_not_overflow(self):
        with self.assertRaisesRegex(ValueError, "estimated duration exceeds"):
            plan_script({"lines": [{"text": "hello", "character": "Reimu"}],
                         "chars_per_sec": 5e-324})

    def test_small_valid_speed_is_accepted(self):
        result = plan_script({"lines": [{"text": "a", "character": "Reimu"}],
                              "chars_per_sec": 0.001, "fps": 30})
        self.assertEqual(result["details"][0]["length"], 30000)
        self.assertTrue(result["dry_run"])

    def test_boundary_exceeding_frame_range_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "estimated duration exceeds"):
            plan_script({"lines": [{"text": "a", "character": "Reimu"}],
                         "chars_per_sec": 1e-12, "fps": 240})


if __name__ == "__main__":
    unittest.main()
