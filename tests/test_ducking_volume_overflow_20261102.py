"""Regression: individually finite volumes can overflow when multiplied."""
import math
import pathlib
import sys
import unittest

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1] / "mcp-server"))
from ducking import plan


class DuckingVolumeOverflowTests(unittest.TestCase):
    def test_overflowed_low_volume_is_rejected(self):
        bgm = {"frame": 0, "length": 100}
        voices = [{"frame": 20, "length": 10}]
        with self.assertRaisesRegex(ValueError, "ducked volume must be finite"):
            plan(bgm, voices, 1.7e308, ratio=0.99)

    def test_large_finite_volume_stays_finite(self):
        bgm = {"frame": 0, "length": 100}
        voices = [{"frame": 20, "length": 10}]
        points = plan(bgm, voices, 1e308, ratio=0.5)
        self.assertTrue(points)
        self.assertTrue(all(math.isfinite(p["value"]) for p in points))


if __name__ == "__main__":
    unittest.main()
