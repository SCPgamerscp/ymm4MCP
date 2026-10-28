"""Preview QA must not certify frames or audio that changed during analysis."""
import os
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock, patch

import gemini_video


class GeminiPreviewSnapshotIntegrityTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        self.frame = self.folder / "0001.png"
        self.frame.write_bytes(b"frame-original")
        self.audio = self.folder / "audio.wav"
        self.audio.write_bytes(b"RIFF\x04\x00\x00\x00WAVE")
        self.client = MagicMock()
        self.client.models.generate_content.return_value = SimpleNamespace(
            text='{"summary":"ok","events":[]}')
        self.part = SimpleNamespace(
            from_bytes=lambda **kw: SimpleNamespace(inline_data=kw["data"]),
            from_text=lambda **kw: SimpleNamespace(inline_data=None))

    def analyze(self, **kwargs):
        with patch.object(gemini_video, "is_available", return_value=(True, "ok")), \
             patch.object(gemini_video, "_client", return_value=self.client), \
             patch.object(gemini_video, "genai_types", SimpleNamespace(Part=self.part)):
            return gemini_video.analyze_frames_dir(str(self.folder), **kwargs)

    def test_unchanged_frame_and_audio_pass(self):
        result = self.analyze(audio_path=str(self.audio))
        self.assertTrue(result["success"])
        self.client.models.generate_content.assert_called_once()

    def test_frame_replaced_during_model_call_fails_closed(self):
        def mutate(**kwargs):
            self.frame.write_bytes(b"changed")
            return SimpleNamespace(text='{"summary":"ok","events":[]}')
        self.client.models.generate_content.side_effect = mutate
        result = self.analyze()
        self.assertFalse(result["success"])
        self.assertEqual(result["source"], "frames_dir")
        self.assertIn("changed during", result["error"])

    def test_audio_replaced_during_model_call_fails_closed(self):
        def mutate(**kwargs):
            self.audio.write_bytes(b"RIFF\x04\x00\x00\x00WAVX")
            return SimpleNamespace(text='{"summary":"ok","events":[]}')
        self.client.models.generate_content.side_effect = mutate
        result = self.analyze(audio_path=str(self.audio))
        self.assertFalse(result["success"])
        self.assertIn("changed during", result["error"])

    def test_same_size_rewrite_with_restored_timestamp_is_detected(self):
        old_stat = self.frame.stat()
        def mutate(**kwargs):
            self.frame.write_bytes(b"frame-modified")
            os.utime(self.frame, ns=(old_stat.st_atime_ns, old_stat.st_mtime_ns))
            return SimpleNamespace(text='{"summary":"ok","events":[]}')
        self.assertEqual(len(b"frame-original"), len(b"frame-modified"))
        self.client.models.generate_content.side_effect = mutate
        result = self.analyze()
        self.assertFalse(result["success"])
        self.assertIn("changed during", result["error"])

    def test_deleted_frame_during_model_call_is_not_success(self):
        def mutate(**kwargs):
            self.frame.unlink()
            return SimpleNamespace(text='{"summary":"ok","events":[]}')
        self.client.models.generate_content.side_effect = mutate
        result = self.analyze()
        self.assertFalse(result["success"])
        self.assertIn("changed during", result["error"])

    def test_frame_changed_while_building_request_never_calls_model(self):
        def from_bytes(**kwargs):
            self.frame.write_bytes(b"frame-changed")
            return SimpleNamespace(inline_data=kwargs["data"])
        self.part.from_bytes = from_bytes
        result = self.analyze()
        self.assertFalse(result["success"])
        self.assertIn("before Gemini", result["error"])
        self.client.models.generate_content.assert_not_called()


if __name__ == "__main__":
    unittest.main()
