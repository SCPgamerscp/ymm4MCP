"""Regression tests for ducking volume precision guards."""
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "mcp-server"))
from ducking import plan


class DuckingPrecisionTests(unittest.TestCase):
    BGM = {"frame": 0, "length": 100}
    VOICES = [{"frame": 20, "length": 10}]

    def test_ratio_that_rounds_to_same_volume_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "does not reduce volume"):
            plan(self.BGM, self.VOICES, 1e-323, ratio=0.9)

    def test_representable_attenuation_is_allowed(self):
        points = plan(self.BGM, self.VOICES, 1.0, ratio=0.3)
        self.assertIn({"at": 20, "value": 0.3}, points)

    def test_underflow_to_zero_still_rejected(self):
        with self.assertRaisesRegex(ValueError, "underflows to zero"):
            plan(self.BGM, self.VOICES, 5e-324, ratio=0.5)


if __name__ == "__main__":
    unittest.main()
