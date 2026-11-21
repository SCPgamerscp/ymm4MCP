"""Regression: empty rendered WAV must not pass audio QA (#61)."""
import os
import tempfile
import unittest
import wave

from wav_audio_qa import inspect_wav


class EmptyWavRegressionTests(unittest.TestCase):
    def test_zero_frame_wav_is_rejected(self):
        with tempfile.TemporaryDirectory() as directory:
            path = os.path.join(directory, "empty.wav")
            with wave.open(path, "wb") as wav:
                wav.setnchannels(2)
                wav.setsampwidth(2)
                wav.setframerate(48000)
                wav.writeframes(b"")
            with self.assertRaisesRegex(ValueError, "no audio frames"):
                inspect_wav(path)


if __name__ == "__main__":
    unittest.main()
