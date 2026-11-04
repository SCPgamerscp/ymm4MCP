"""Regression tests for bounded WAV audio inspection."""
import os
import tempfile
import unittest
import wave

from wav_audio_qa import inspect_wav


def write_wav(path, samples, rate=10, channels=1, width=2):
    with wave.open(path, "wb") as output:
        output.setnchannels(channels)
        output.setsampwidth(width)
        output.setframerate(rate)
        output.writeframes(b"".join(int(v).to_bytes(width, "little", signed=True)
                                    for v in samples))


class AudioQaTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = os.path.join(self.temp.name, "test.wav")

    def test_detects_long_silence_and_clipping(self):
        write_wav(self.path, [0] * 20 + [32767] * 10)
        result = inspect_wav(self.path)
        self.assertFalse(result["passed"])
        self.assertEqual(result["clipped_samples"], 10)
        self.assertEqual(result["long_silences"], [{"start_seconds": 0.0, "end_seconds": 2.0}])

    def test_clean_audio_passes(self):
        write_wav(self.path, [10000] * 30)
        self.assertTrue(inspect_wav(self.path)["passed"])

    def test_stereo_requires_all_channels_quiet(self):
        write_wav(self.path, [0, 10000] * 20, channels=2)
        self.assertEqual(inspect_wav(self.path)["long_silences"], [])

    def test_rejects_oversized_input_and_bad_thresholds(self):
        write_wav(self.path, [0] * 10)
        with self.assertRaises(ValueError):
            inspect_wav(self.path, max_bytes=1)
        with self.assertRaises(ValueError):
            inspect_wav(self.path, clip_threshold=float("nan"))

    def test_truncated_pcm_is_not_success(self):
        write_wav(self.path, [1000] * 20)
        with open(self.path, "rb+") as stream:
            stream.truncate(os.path.getsize(self.path) - 10)
        with self.assertRaises((ValueError, EOFError)):
            inspect_wav(self.path)


if __name__ == "__main__":
    unittest.main()
