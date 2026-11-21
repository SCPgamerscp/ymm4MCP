"""Regression tests for optional per-channel DC offset gate."""
import os
import tempfile
import unittest
import wave

from wav_audio_qa import inspect_wav


def write_pcm(path, values, channels=1):
    with wave.open(path, "wb") as output:
        output.setnchannels(channels)
        output.setsampwidth(2)
        output.setframerate(100)
        output.writeframes(b"".join(int(v).to_bytes(2, "little", signed=True)
                                    for v in values))


class DcOffsetTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.path = os.path.join(self.tmp.name, "test.wav")

    def test_constant_offset_fails_when_gate_enabled(self):
        write_pcm(self.path, [8000] * 200)
        result = inspect_wav(self.path, max_dc_offset_dbfs=-20)
        self.assertFalse(result["passed"])
        self.assertEqual(result["issues"][0]["code"], "AUDIO_DC_OFFSET")
        self.assertEqual(result["issues"][0]["channels"], [0])

    def test_zero_mean_wave_passes(self):
        write_pcm(self.path, [8000, -8000] * 100)
        result = inspect_wav(self.path, max_dc_offset_dbfs=-20)
        self.assertTrue(result["passed"])
        self.assertEqual(result["channel_dc_offset_dbfs"], [None])

    def test_stereo_gate_identifies_only_offset_channel(self):
        write_pcm(self.path, [8000, 8000, 8000, -8000] * 100, channels=2)
        result = inspect_wav(self.path, max_dc_offset_dbfs=-20)
        self.assertEqual(result["issues"][0]["channels"], [0])
        self.assertIsNone(result["channel_dc_offset_dbfs"][1])

    def test_default_does_not_reject_dc_offset(self):
        write_pcm(self.path, [8000] * 200)
        self.assertTrue(inspect_wav(self.path)["passed"])

    def test_invalid_dc_threshold_rejected(self):
        write_pcm(self.path, [8000] * 200)
        for value in (float("nan"), float("inf"), True, -121, 1, "bad"):
            with self.subTest(value=value), self.assertRaises(ValueError):
                inspect_wav(self.path, max_dc_offset_dbfs=value)


if __name__ == "__main__":
    unittest.main()
