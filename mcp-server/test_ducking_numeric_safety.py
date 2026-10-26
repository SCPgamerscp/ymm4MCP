"""Numerical edge cases must not escape BGM ducking validation."""
import unittest

import ducking


BGM = {"frame": 0, "length": 100}
VOICES = [{"frame": 20, "length": 10}]


class DuckingNumericSafetyTests(unittest.TestCase):
    def test_oversized_integer_ratio_returns_input_error(self):
        with self.assertRaisesRegex(ValueError, "duck_ratio must be a finite number"):
            ducking.plan(BGM, VOICES, 100.0, ratio=10 ** 400)

    def test_oversized_integer_volume_returns_input_error(self):
        with self.assertRaisesRegex(ValueError, "base_volume must be a finite number"):
            ducking.plan(BGM, VOICES, 10 ** 400)

    def test_tiny_ratio_must_not_silently_mute_audio(self):
        with self.assertRaisesRegex(ValueError, "underflows to zero"):
            ducking.plan(BGM, VOICES, 1.0, ratio=5e-324)

    def test_tiny_volume_must_not_silently_mute_audio(self):
        with self.assertRaisesRegex(ValueError, "underflows to zero"):
            ducking.plan(BGM, VOICES, 5e-324, ratio=0.3)

    def test_normal_curve_is_unchanged(self):
        self.assertEqual(ducking.plan(BGM, VOICES, 100.0), [
            {"at": 15, "value": 100.0},
            {"at": 20, "value": 30.0},
            {"at": 30, "value": 30.0},
            {"at": 40, "value": 100.0},
        ])


if __name__ == "__main__":
    unittest.main()
