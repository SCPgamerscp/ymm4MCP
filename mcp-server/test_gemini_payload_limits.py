"""Gemini frame analysis must bound file reads and total inline payloads."""
import tempfile
from pathlib import Path
from types import SimpleNamespace
import unittest
from unittest.mock import MagicMock, patch

import gemini_video


class GeminiPayloadLimitTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        self.frame = self.folder / "0001.png"
        self.frame.write_bytes(b"png")
        self.audio = self.folder / "audio.wav"
        self.client = MagicMock()
        self.client.models.generate_content.return_value = SimpleNamespace(
            text='{"summary":"ok","events":[]}')
        self.part = SimpleNamespace(
            from_bytes=lambda **kwargs: SimpleNamespace(inline_data=kwargs["data"]),
            from_text=lambda **kwargs: SimpleNamespace(inline_data=None),
        )

    def analyze(self, **kwargs):
        with patch.object(gemini_video, "is_available", return_value=(True, "ok")), \
             patch.object(gemini_video, "_client", return_value=self.client), \
             patch.object(gemini_video, "genai_types", SimpleNamespace(Part=self.part)):
            return gemini_video.analyze_frames_dir(str(self.folder), **kwargs)

    def test_large_frame_is_rejected_before_upload(self):
        self.frame.write_bytes(b"12345")
        with patch.object(gemini_video, "MAX_FRAME_IMAGE_BYTES", 4):
            result = self.analyze()
        self.assertFalse(result["success"])
        self.assertIn("limit", result["error"])
        self.client.models.generate_content.assert_not_called()

    def test_large_audio_is_rejected_before_upload(self):
        self.audio.write_bytes(b"12345")
        with patch.object(gemini_video, "MAX_AUDIO_BYTES", 4):
            result = self.analyze(audio_path=str(self.audio))
        self.assertFalse(result["success"])
        self.client.models.generate_content.assert_not_called()

    def test_total_frame_audio_budget_is_enforced(self):
        self.frame.write_bytes(b"1234")
        self.audio.write_bytes(b"abcd")
        with patch.object(gemini_video, "MAX_TOTAL_BYTES", 7):
            result = self.analyze(audio_path=str(self.audio))
        self.assertFalse(result["success"])
        self.client.models.generate_content.assert_not_called()

    def test_small_payload_still_succeeds(self):
        import wave
        with wave.open(str(self.audio), "wb") as wav:
            wav.setnchannels(1)
            wav.setsampwidth(2)
            wav.setframerate(8000)
            wav.writeframes(bytes(2))
        result = self.analyze(audio_path=str(self.audio))
        self.assertTrue(result["success"])
        self.assertEqual(result["analyzed_images"], 1)
        self.client.models.generate_content.assert_called_once()


if __name__ == "__main__":
    unittest.main()
