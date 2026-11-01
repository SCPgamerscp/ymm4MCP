"""Regression tests for malformed host mutation expectations."""
import pathlib
import sys
import unittest

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1] / "mcp-server"))
from ducking import verify_mutation_point


class MutationEchoGuardTests(unittest.TestCase):
    def test_malformed_expectations_fail_closed(self):
        echo = [{"at": 3, "value": 0.3}]
        for point in (None, [], {}, {"at": 3}, {"value": 0.3},
                      {"at": True, "value": 0.3},
                      {"at": 3, "value": True},
                      {"at": 3, "value": "0.3"}):
            with self.subTest(point=point):
                self.assertFalse(verify_mutation_point(echo, point))

    def test_valid_echo_and_conflicting_echo(self):
        point = {"at": 3, "value": 0.3}
        self.assertTrue(verify_mutation_point([point.copy()], point))
        self.assertFalse(verify_mutation_point([point.copy(), point.copy()], point))
        self.assertFalse(verify_mutation_point([{"at": 3, "value": 0.4}], point))


if __name__ == "__main__":
    unittest.main()
