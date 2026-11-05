"""Regression checks for clipped-sample fraction (#61)."""
import importlib.util
from pathlib import Path
import struct
import tempfile
import unittest
import wave

MODULE_PATH = Path(__file__).with_name("wav_audio_qa.py")
spec = importlib.util.spec_from_file_location("wav_audio_qa", MODULE_PATH)
audio_qa = importlib.util.module_from_spec(spec)
spec.loader.exec_module(audio_qa)


class ClippedSampleFractionTests(unittest.TestCase):
    def test_fraction_counts_samples_across_channels(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "stereo.wav"
            with wave.open(str(path), "wb") as output:
                output.setnchannels(2)
                output.setsampwidth(2)
                output.setframerate(4)
                output.writeframes(struct.pack("<hhhhhhhh", 32767, 1000, 1000, 1000,
                                               -32768, 1000, 1000, 1000))
            result = audio_qa.inspect_wav(path, min_silence_seconds=2)
            self.assertEqual(result["clipped_samples"], 2)
            self.assertAlmostEqual(result["clipped_sample_fraction"], 0.25)
            self.assertFalse(result["passed"])

    def test_clean_audio_has_zero_fraction(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "mono.wav"
            with wave.open(str(path), "wb") as output:
                output.setnchannels(1)
                output.setsampwidth(2)
                output.setframerate(4)
                output.writeframes(struct.pack("<hhhh", 1000, -1000, 1000, -1000))
            result = audio_qa.inspect_wav(path, min_silence_seconds=2)
            self.assertEqual(result["clipped_sample_fraction"], 0.0)
            self.assertTrue(result["passed"])


if __name__ == "__main__":
    unittest.main()
