"""Gemini frame QA must use chronological samples and reject incomplete media."""
import tempfile
from pathlib import Path
from types import SimpleNamespace
import unittest
from unittest.mock import MagicMock, patch

import gemini_video


class GeminiFrameOrderIntegrityTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        for number in (1, 2, 10):
            (self.folder / f"{number}.png").write_bytes(f"frame-{number}".encode())
        self.images = []
        self.labels = []
        self.client = MagicMock()
        self.client.models.generate_content.return_value = SimpleNamespace(
            text='{"summary":"ok","events":[]}')
        self.part = SimpleNamespace(
            from_bytes=self._bytes_part,
            from_text=self._text_part,
        )

    def _bytes_part(self, *, data, mime_type):
        if mime_type == "image/png":
            self.images.append(data)
        return SimpleNamespace(inline_data=data)

    def _text_part(self, *, text):
        self.labels.append(text)
        return SimpleNamespace(inline_data=None)

    def analyze(self, **kwargs):
        with patch.object(gemini_video, "is_available", return_value=(True, "ok")), \
             patch.object(gemini_video, "_client", return_value=self.client), \
             patch.object(gemini_video, "genai_types", SimpleNamespace(Part=self.part)):
            return gemini_video.analyze_frames_dir(str(self.folder), **kwargs)

    def test_shuffled_metadata_is_sorted_by_capture_time(self):
        metadata = [
            {"file": "10.png", "timeMs":10000, "frame":300},
            {"file": "1.png", "timeMs":1000, "frame":30},
            {"file": "2.png", "timeMs":2000, "frame":60},
        ]
        result = self.analyze(frames_meta=metadata)
        self.assertTrue(result["success"])
        self.assertEqual(self.images, [b"frame-1", b"frame-2", b"frame-10"])
        self.assertIn("1.00s", self.labels[0])
        self.assertIn("10.00s", self.labels[2])

    def test_frame_indices_are_fallback_when_times_absent(self):
        metadata = [
            {"file": "10.png", "frame":10},
            {"file": "2.png", "frame":2},
            {"file": "1.png", "frame":1},
        ]
        self.assertTrue(self.analyze(frames_meta=metadata)["success"])
        self.assertEqual(self.images, [b"frame-1", b"frame-2", b"frame-10"])

    def test_numeric_filenames_are_naturally_ordered(self):
        self.assertTrue(self.analyze()["success"])
        self.assertEqual(self.images, [b"frame-1", b"frame-2", b"frame-10"])

    def test_incomparable_partial_timing_metadata_is_rejected(self):
        result = self.analyze(frames_meta=[
            {"file": "1.png", "timeMs":1000},
            {"file": "2.png", "timeMs":None},
        ])
        self.assertFalse(result["success"])
        self.assertIn("inconsistent", result["error"])
        self.client.models.generate_content.assert_not_called()

    def test_missing_selected_frame_is_not_silently_ignored(self):
        result = self.analyze(frames_meta=[
            {"file": "1.png", "timeMs":1000},
            {"file": "missing.png", "timeMs":2000},
        ])
        self.assertFalse(result["success"])
        self.assertEqual(result["source"], "frames_dir")
        self.assertIn("missing", result["error"])
        self.client.models.generate_content.assert_not_called()

    def test_requested_missing_audio_is_not_silently_omitted(self):
        result = self.analyze(audio_path=str(self.folder / "missing.wav"))
        self.assertFalse(result["success"])
        self.assertEqual(result["source"], "frames_dir")
        self.assertIn("audio", result["error"])
        self.client.models.generate_content.assert_not_called()

    def test_audio_is_optional_when_not_requested(self):
        self.assertTrue(self.analyze()["success"])
        self.client.models.generate_content.assert_called_once()

    def test_wrong_audio_extension_is_rejected(self):
        bad = self.folder / "sound.mp3"
        bad.write_bytes(b"ID3")
        result = self.analyze(audio_path=str(bad))
        self.assertFalse(result["success"])
        self.assertIn("WAV", result["error"])
        self.client.models.generate_content.assert_not_called()

    def test_wrong_wav_signature_is_rejected(self):
        bad = self.folder / "sound.wav"
        bad.write_bytes(b"ID3not-a-wav-file")
        result = self.analyze(audio_path=str(bad))
        self.assertFalse(result["success"])
        self.assertIn("RIFF/WAVE", result["error"])
        self.client.models.generate_content.assert_not_called()

    def test_valid_wav_header_is_accepted(self):
        audio = self.folder / "sound.wav"
        audio.write_bytes(b"RIFF\x04\x00\x00\x00WAVE")
        result = self.analyze(audio_path=str(audio))
        self.assertTrue(result["success"])
        self.client.models.generate_content.assert_called_once()


if __name__ == "__main__":
    unittest.main()
