"""Regression tests for Gemini frame export containment and input validation."""
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import gemini_video


class GeminiFramePathTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name) / "frames"
        self.root.mkdir()
        self.inside = self.root / "0001.png"
        self.inside.write_bytes(b"image")

    def test_contained_png_is_allowed(self):
        self.assertEqual(gemini_video._safe_frame_path(str(self.root), "0001.png"),
                         str(self.inside.resolve()))

    def test_absolute_and_parent_paths_are_rejected(self):
        outside = Path(self.temp.name) / "secret.png"
        outside.write_bytes(b"secret")
        for name in ("../secret.png", str(outside), "0001.txt", 1, None):
            with self.subTest(name=name):
                self.assertIsNone(gemini_video._safe_frame_path(str(self.root), name))

    def test_symlink_outside_export_is_rejected(self):
        outside = Path(self.temp.name) / "secret.png"
        outside.write_bytes(b"secret")
        link = self.root / "linked.png"
        try:
            link.symlink_to(outside)
        except (OSError, NotImplementedError):
            self.skipTest("symlinks unavailable on this platform")
        self.assertIsNone(gemini_video._safe_frame_path(str(self.root), link.name))

    def test_invalid_metadata_and_image_limit_fail_cleanly(self):
        with patch.object(gemini_video, "is_available", return_value=(True, "ok")):
            for kwargs in ({"max_images": 0}, {"max_images": True},
                           {"frames_meta": ["not an object"]}):
                with self.subTest(kwargs=kwargs):
                    result = gemini_video.analyze_frames_dir(str(self.root), **kwargs)
                    self.assertFalse(result["success"])

    def test_outside_only_metadata_does_not_call_gemini(self):
        outside = Path(self.temp.name) / "secret.png"
        outside.write_bytes(b"secret")
        with patch.object(gemini_video, "is_available", return_value=(True, "ok")), \
             patch.object(gemini_video, "_client") as client:
            result = gemini_video.analyze_frames_dir(
                str(self.root), frames_meta=[{"file": "../secret.png"}])
        self.assertFalse(result["success"])
        self.assertEqual(result["source"], "frames_dir")
        client.return_value.models.generate_content.assert_not_called()


if __name__ == "__main__":
    unittest.main()
