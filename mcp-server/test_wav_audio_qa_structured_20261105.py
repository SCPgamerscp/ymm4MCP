"""Regression tests for structured WAV QA and RMS stereo imbalance (#61)."""
import math
import os
import tempfile
import unittest
import wave

from wav_audio_qa import inspect_wav


def make_stereo(path, left, right):
    with wave.open(path, "wb") as output:
        output.setnchannels(2)
        output.setsampwidth(2)
        output.setframerate(10)
        output.writeframes(b"".join(
            int(sample).to_bytes(2, "little", signed=True)
            for pair in zip(left, right) for sample in pair))


class StructuredAudioQaTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = os.path.join(self.temp.name, "stereo.wav")

    def test_balanced_audio_passes(self):
        make_stereo(self.path, [10000] * 30, [-10000] * 30)
        result = inspect_wav(self.path, max_stereo_imbalance_db=3)
        self.assertTrue(result["passed"])
        self.assertEqual(result["issues"], [])
        self.assertAlmostEqual(result["stereo_imbalance_db"], 0)

    def test_imbalance_is_a_failing_structured_finding(self):
        make_stereo(self.path, [10000] * 30, [1000] * 30)
        result = inspect_wav(self.path, max_stereo_imbalance_db=6)
        self.assertFalse(result["passed"])
        self.assertAlmostEqual(result["stereo_imbalance_db"], 20)
        self.assertEqual([item["code"] for item in result["issues"]],
                         ["AUDIO_STEREO_IMBALANCE"])

    def test_imbalance_is_opt_in(self):
        make_stereo(self.path, [10000] * 30, [1000] * 30)
        self.assertTrue(inspect_wav(self.path)["passed"])

    def test_silent_channel_never_passes_imbalance_gate(self):
        make_stereo(self.path, [10000] * 30, [0] * 30)
        result = inspect_wav(self.path, max_stereo_imbalance_db=120)
        self.assertIsNone(result["stereo_imbalance_db"])
        self.assertEqual([issue["code"] for issue in result["issues"]],
                         ["AUDIO_SILENT_CHANNEL", "AUDIO_STEREO_IMBALANCE"])

    def test_both_channels_silent_not_infinite_imbalance(self):
        make_stereo(self.path, [0] * 30, [0] * 30)
        result = inspect_wav(self.path, max_stereo_imbalance_db=3)
        self.assertEqual(result["stereo_imbalance_db"], 0)
        self.assertNotIn("AUDIO_STEREO_IMBALANCE",
                         [issue["code"] for issue in result["issues"]])
        self.assertIn("AUDIO_LONG_SILENCE",
                      [issue["code"] for issue in result["issues"]])

    def test_bad_imbalance_threshold_rejected(self):
        make_stereo(self.path, [10000] * 30, [10000] * 30)
        for invalid in (float("nan"), float("inf"), -1, 121, True, "6"):
            with self.subTest(invalid=invalid):
                with self.assertRaises(ValueError):
                    inspect_wav(self.path, max_stereo_imbalance_db=invalid)


if __name__ == "__main__":
    unittest.main()
