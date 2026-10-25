"""Regression tests for fail-closed BGM ducking input and curve verification."""

import unittest

import ducking


class DuckingValidationTests(unittest.TestCase):
    def setUp(self):
        self.bgm = {"frame": 0, "length": 100}
        self.voices = [{"frame": 20, "length": 10}]
        self.points = [{"at": 20, "value": 30.0}]
        self.report = {
            "success": True,
            "revision": "r2",
            "animations": [{"prop": "Volume", "keyframes": [
                {"at": 0, "value": 100.0},
                {"at": 20, "value": 30.0},
            ]}],
        }

    def test_valid_curve_still_verifies(self):
        self.assertTrue(ducking.verify_keyframes(self.report, "r2", self.points, 100))

    def test_nonfinite_expected_value_never_verifies(self):
        for invalid in (float("nan"), float("inf"), float("-inf")):
            with self.subTest(value=invalid):
                self.assertFalse(ducking.verify_keyframes(
                    self.report, "r2", [{"at": 20, "value": invalid}], 100))

    def test_nonfinite_base_volume_never_verifies(self):
        for invalid in (float("nan"), float("inf"), float("-inf")):
            with self.subTest(value=invalid):
                self.assertFalse(ducking.verify_keyframes(
                    self.report, "r2", self.points, invalid))

    def test_duplicate_or_malformed_expected_points_never_verify(self):
        for invalid in (
            [{"at": 20, "value": 30.0}, {"at": 20, "value": 30.0}],
            [{"at": True, "value": 30.0}],
            [{"at": 20, "value": True}],
            [{"at": 20, "value": "30"}],
            [{"at": -1, "value": 30.0}],
            [{"at": 20}],
            [None],
            None,
        ):
            with self.subTest(points=invalid):
                self.assertFalse(ducking.verify_keyframes(
                    self.report, "r2", invalid, 100))

    def test_malformed_bgm_and_voice_objects_raise_value_error(self):
        for invalid in (None, "audio", []):
            with self.subTest(bgm=invalid), self.assertRaises(ValueError):
                ducking.plan(invalid, self.voices, 100)
        for invalid in (None, "voice", []):
            with self.subTest(voice=invalid), self.assertRaises(ValueError):
                ducking.plan(self.bgm, [invalid], 100)


if __name__ == "__main__":
    unittest.main()
