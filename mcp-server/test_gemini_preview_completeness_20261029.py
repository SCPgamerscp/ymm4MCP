"""Regression tests for complete Gemini preview sampling (#61)."""
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock, patch

import gemini_video


class GeminiPreviewCompletenessTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        self.available = patch.object(gemini_video, "is_available", return_value=(True, "ok"))
        self.available.start()
        self.addCleanup(self.available.stop)

    def test_explicit_empty_capture_does_not_scan_stale_pngs(self):
        (self.folder / "old.png").write_bytes(b"old frame")
        with patch.object(gemini_video, "_client") as client:
            result = gemini_video.analyze_frames_dir(str(self.folder), frames_meta=[])
        self.assertFalse(result["success"], result)
        self.assertIn("no captured frames", result["error"])
        client.assert_not_called()

    def test_downsampling_keeps_first_middle_and_final_frame(self):
        for i in range(5):
            (self.folder / f"{i}.png").write_bytes(f"image-{i}".encode())
        sent = []
        class Part:
            @staticmethod
            def from_bytes(*, data, mime_type):
                sent.append(data)
                return data

            @staticmethod
            def from_text(*, text):
                return text

        client = MagicMock()
        client.models.generate_content.return_value = SimpleNamespace(
            text='{"summary":"ok","events":[]}')
        with patch.object(gemini_video, "_client", return_value=client), \\
             patch.object(gemini_video, "genai_types", SimpleNamespace(Part=Part)):
            result = gemini_video.analyze_frames_dir(str(self.folder), max_images=3)
        self.assertTrue(result["success"], result)
        self.assertEqual(sent, [b"image-0", b"image-2", b"image-4"])
        self.assertEqual(result["analyzed_images"], 3)

    def test_single_image_limit_remains_supported(self):
        for i in range(3):
            (self.folder / f"{i}.png").write_bytes(f"image-{i}".encode())
        sent = []
        class Part:
            @staticmethod
            def from_bytes(*, data, mime_type):
                sent.append(data)
                return data

            @staticmethod
            def from_text(*, text):
                return text

        client = MagicMock()
        client.models.generate_content.return_value = SimpleNamespace(
            text='{"summary":"ok","events":[]}')
        with patch.object(gemini_video, "_client", return_value=client), \\
             patch.object(gemini_video, "genai_types", SimpleNamespace(Part=Part)):
            result = gemini_video.analyze_frames_dir(str(self.folder), max_images=1)
        self.assertTrue(result["success"], result)
        self.assertEqual(sent, [b"image-0"])

    def test_zero_byte_selected_png_cannot_pass_visual_qa(self):
        (self.folder / "0.png").write_bytes(b"")
        with patch.object(gemini_video, "_client") as client:
            result = gemini_video.analyze_frames_dir(str(self.folder))
        self.assertFalse(result["success"], result)
        self.assertIn("empty", result["error"])
        client.return_value.models.generate_content.assert_not_called()


if __name__ == "__main__":
    unittest.main()
