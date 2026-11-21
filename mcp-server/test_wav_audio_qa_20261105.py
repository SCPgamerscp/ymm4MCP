"""Regression tests for bounded WAV audio inspection."""
import math
import os
import struct
import tempfile
import unittest
import wave

from wav_audio_qa import inspect_wav


def write_wav(path, samples, rate=10, channels=1, width=2):
    with wave.open(path, "wb") as output:
        output.setnchannels(channels)
        output.setsampwidth(width)
        output.setframerate(rate)
        if width == 1:  # 8-bit PCM uses unsigned storage.
            output.writeframes(bytes(v + 128 for v in samples))
        else:
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

    def test_pcm_rails_clip_in_both_directions(self):
        for width in (1, 2, 3, 4):
            with self.subTest(width=width):
                scale = 1 << (8 * width - 1)
                write_wav(self.path, [scale - 1, scale - 2, -scale, -scale + 1], width=width)
                result = inspect_wav(self.path, clip_threshold=1.0)
                self.assertEqual(result["clipped_samples"], 2)
                self.assertEqual(result["clipped_sample_fraction"], 0.5)

    def test_default_threshold_detects_unsigned_8_bit_clipping(self):
        write_wav(self.path, [127, 126], width=1)
        self.assertEqual(inspect_wav(self.path)["clipped_samples"], 1)

    def test_stereo_requires_all_channels_quiet(self):
        write_wav(self.path, [0, 10000] * 20, channels=2)
        self.assertEqual(inspect_wav(self.path)["long_silences"], [])

    def test_silent_left_channel_fails_stereo_qa(self):
        write_wav(self.path, [0, 10000] * 20, channels=2)
        result = inspect_wav(self.path)
        self.assertEqual(result["silent_channels"], [0])
        self.assertIsNone(result["channel_peak_dbfs"][0])
        self.assertIsNotNone(result["channel_peak_dbfs"][1])
        self.assertFalse(result["passed"])

    def test_active_stereo_channels_pass(self):
        write_wav(self.path, [10000, -12000] * 20, channels=2)
        result = inspect_wav(self.path)
        self.assertEqual(result["silent_channels"], [])
        self.assertEqual(len(result["channel_peak_dbfs"]), 2)
        self.assertTrue(result["passed"])

    def test_rms_reports_constant_mono_amplitude(self):
        write_wav(self.path, [10000] * 30)
        result = inspect_wav(self.path)
        expected = 20 * math.log10(10000 / 32768)
        self.assertAlmostEqual(result["rms_dbfs"], expected)
        self.assertAlmostEqual(result["channel_rms_dbfs"][0], expected)
        self.assertTrue(result["passed"])

    def test_rms_exposes_silent_stereo_channel(self):
        write_wav(self.path, [0, 10000] * 20, channels=2)
        result = inspect_wav(self.path)
        self.assertEqual(result["channel_rms_dbfs"][0], None)
        self.assertAlmostEqual(result["channel_rms_dbfs"][1], 20 * math.log10(10000 / 32768))
        self.assertAlmostEqual(result["rms_dbfs"], 20 * math.log10(10000 / 32768 / math.sqrt(2)))

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

    def test_rejects_partial_pcm_frame_at_end_of_data(self):
        for tail in (b"", b"\x01"):
            with self.subTest(tail=tail):
                write_wav(self.path, [1000] * 4, channels=2)
                with open(self.path, "rb+") as stream:
                    stream.seek(4)
                    stream.write(struct.pack("<I", 46))  # 36 header bytes, 9 data bytes, 1 pad byte.
                    stream.seek(40)
                    stream.write(struct.pack("<I", 9))
                    stream.seek(0, os.SEEK_END)
                    if tail:
                        stream.write(tail + b"\x00")
                with self.assertRaises(ValueError):
                    inspect_wav(self.path)

    def test_accepts_complete_odd_byte_pcm_data(self):
        write_wav(self.path, [50, 60, 70], width=1)
        self.assertEqual(inspect_wav(self.path)["duration_seconds"], 0.3)


if __name__ == "__main__":
    unittest.main()
