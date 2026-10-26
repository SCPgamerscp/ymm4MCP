"""Ducking verification must never accept unversioned or invalid baselines."""

import unittest

import ducking


class DuckingRevisionGuardTests(unittest.TestCase):
    def setUp(self):
        self.report = {
            "success": True,
            "revision": "r2",
            "animations": [{"prop": "Volume", "keyframes": [
                {"at": 0, "value": 100.0},
                {"at": 20, "value": 30.0},
            ]}],
        }
        self.points = [{"at": 20, "value": 30.0}]

    def test_valid_versioned_curve_still_verifies(self):
        self.assertTrue(ducking.verify_keyframes(self.report, "r2", self.points, 100.0))

    def test_missing_or_blank_revision_cannot_verify(self):
        for revision in (None, "", 0):
            report = {**self.report, "revision": revision}
            with self.subTest(revision=revision):
                self.assertFalse(ducking.verify_keyframes(report, revision, self.points, 100.0))

    def test_nonpositive_baseline_cannot_verify(self):
        for baseline in (0, -1, -100):
            report = {**self.report, "animations": [{"prop": "Volume", "keyframes": [
                {"at": 0, "value": baseline}, {"at": 20, "value": 30.0}]}]}
            with self.subTest(baseline=baseline):
                self.assertFalse(ducking.verify_keyframes(report, "r2", self.points, baseline))


if __name__ == "__main__":
    unittest.main()
