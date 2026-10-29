"""Gemini video analysis should fail cleanly on bad paths and sanitize frame labels."""
import tempfile
from pathlib import Path
from types import SimpleNamespace
import unittest
from unittest.mock import MagicMock, patch

import gemini_video


class GeminiFrameMetadataTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.frame = Path(self.temp.name) / "0001.png"
        self.frame.write_bytes(b"fake png")
        self.client = MagicMock()
        self.client.models.generate_content.return_value = SimpleNamespace(
            text='{"summary":"ok","events":[]}')

    def analyze(self, metadata):
        labels = []
        def from_text(*, text):
            labels.append(text)
            return SimpleNamespace(inline_data=None)
        part = SimpleNamespace(
            from_bytes=lambda **kwargs: SimpleNamespace(inline_data=kwargs["data"]),
            from_text=from_text,
        )
        with patch.object(gemini_video, "is_available", return_value=(True, "ok")), \
             patch.object(gemini_video, "_client", return_value=self.client), \
             patch.object(gemini_video, "genai_types", SimpleNamespace(Part=part)):
            result = gemini_video.analyze_frames_dir(
                self.temp.name, frames_meta=[{"file": "0001.png", **metadata}])
        return result, labels

    def test_invalid_times_and_frame_labels_fail_before_upload(self):
        # A declared but invalid capture timestamp/frame must not be treated
        # as a complete QA result or silently downgraded to unknown labels.
        for time_ms in (True, -1, "1000", float("nan"), float("inf"), 10 ** 500):
            with self.subTest(time_ms=time_ms):
                result, labels = self.analyze({"timeMs": time_ms, "frame": "bad\nframe"})
                self.assertFalse(result["success"], result)
                self.assertIn("inconsistent", result["error"])
                self.assertEqual(labels, [])
        self.client.models.generate_content.assert_not_called()

    def test_valid_time_and_frame_are_preserved(self):
        result, labels = self.analyze({"timeMs": 1500, "frame": 45})
        self.assertTrue(result["success"])
        self.assertIn("時刻 = 1.50s, frame=45", labels[0])

    def test_invalid_file_paths_fail_without_sdk_calls(self):
        with patch.object(gemini_video, "is_available", return_value=(True, "ok")), \
             patch.object(gemini_video, "_client") as client:
            for path in (None, 12, "", "   ", "C:/clip\n.mp4"):
                with self.subTest(path=path):
                    result = gemini_video.analyze_video_file(path)
                    self.assertFalse(result["success"])
            client.assert_not_called()


if __name__ == "__main__":
    unittest.main()
