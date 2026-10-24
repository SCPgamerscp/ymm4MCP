"""Fail closed when the YMM4 host reports a keyframe-read error."""
import unittest

import ducking


class DuckingHostErrorTests(unittest.TestCase):
    def setUp(self):
        self.valid = {"success": True, "revision": "r5", "animations": [
            {"prop": "Volume", "keyframes": [{"at": 0, "value": 100.0},
                                               {"at": 20, "value": 30.0}]}]}
        self.points = [{"at": 20, "value": 30.0}]

    def test_valid_report_still_verifies(self):
        self.assertTrue(ducking.verify_keyframes(self.valid, "r5", self.points, 100))

    def test_host_error_with_matching_curve_is_rejected(self):
        self.assertFalse(ducking.verify_keyframes(
            {**self.valid, "error": "host unavailable"}, "r5", self.points, 100))

    def test_null_error_marker_is_rejected(self):
        self.assertFalse(ducking.verify_keyframes(
            {**self.valid, "error": None}, "r5", self.points, 100))

    def test_error_code_with_matching_curve_is_rejected(self):
        self.assertFalse(ducking.verify_keyframes(
            {**self.valid, "error_code": "ITEMS_UNAVAILABLE"}, "r5", self.points, 100))

    def test_incorrect_revision_is_rejected(self):
        self.assertFalse(ducking.verify_keyframes(self.valid, "r4", self.points, 100))

    def test_duplicate_key_is_rejected(self):
        broken = {**self.valid, "animations": [{"prop": "Volume", "keyframes": [
            {"at": 20, "value": 30}, {"at": 20, "value": 30}]}]}
        self.assertFalse(ducking.verify_keyframes(broken, "r5", self.points, 100))


if __name__ == "__main__":
    unittest.main()
