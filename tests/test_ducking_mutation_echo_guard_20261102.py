"""Regression tests for malformed host mutation expectations."""
import pathlib
import sys
import unittest

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1] / "mcp-server"))
from ducking import verify_mutation_point, verify_keyframes


class MutationEchoGuardTests(unittest.TestCase):
    def test_malformed_expectations_fail_closed(self):
        echo = [{"at": 3, "value": 0.3}]
        for point in (None, [], {}, {"at": 3}, {"value": 0.3},
                      {"at": True, "value": 0.3},
                      {"at": 3, "value": True},
                      {"at": 3, "value": "0.3"},
                      {"at": -1, "value": 0.3},
                      {"at": 2_147_483_648, "value": 0.3}):
            with self.subTest(point=point):
                self.assertFalse(verify_mutation_point(echo, point))

    def test_valid_echo_and_conflicting_echo(self):
        point = {"at": 3, "value": 0.3}
        self.assertTrue(verify_mutation_point([point.copy()], point))
        self.assertFalse(verify_mutation_point([point.copy(), point.copy()], point))
        self.assertFalse(verify_mutation_point([{"at": 3, "value": 0.4}], point))
        # A string echo is not evidence that the host stored a numeric keyframe.
        self.assertFalse(verify_mutation_point([{"at": 3, "value": "0.3"}], point))
        self.assertFalse(verify_mutation_point([{"at": 3, "value": True}], point))


    def test_full_curve_rejects_host_boolean_values(self):
        report = {"success": True, "revision": "rev1", "animations": [
            {"prop": "Volume", "keyframes": [{"at": 0, "value": True}]}]}
        self.assertFalse(verify_keyframes(report, "rev1", [], 1.0))

    def test_full_curve_rejects_out_of_range_absolute_frame(self):
        report = {"success": True, "revision": "rev1", "animations": [
            {"prop": "Volume", "keyframes": [
                {"at": 0, "value": 1.0, "absolute_frame": 2_147_483_647},
                {"at": 1, "value": 0.3, "absolute_frame": 2_147_483_648}]}]}
        self.assertFalse(verify_keyframes(
            report, "rev1", [{"at": 1, "value": 0.3}], 1.0,
            expected_start_frame=2_147_483_647))
        self.assertFalse(verify_keyframes(
            report, "rev1", [{"at": 1, "value": 0.3}], 1.0,
            expected_start_frame=2_147_483_648))

if __name__ == "__main__":
    unittest.main()
