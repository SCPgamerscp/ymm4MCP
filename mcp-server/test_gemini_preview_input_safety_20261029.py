"""Regression tests for bounded and complete Gemini preview QA (#61)."""
from contextlib import nullcontext
from itertools import repeat
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch
import tempfile
import unittest

import gemini_video


class GeminiPreviewInputSafetyTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        available = patch.object(gemini_video, "is_available", return_value=(True, "ok"))
        available.start()
        self.addCleanup(available.stop)

    def reject(self, **kwargs):
        with patch.object(gemini_video, "_client",
                          side_effect=AssertionError("Gemini must not be called")):
            result = gemini_video.analyze_frames_dir(str(self.root), **kwargs)
        self.assertFalse(result["success"], result)
        return result["error"]

    def test_reject_ambiguous_metadata_paths_before_upload(self):
        (self.root / "a.png").write_bytes(b"fixture")
        invalid = ("./a.png", "sub/a.png", "sub\\\\a.png",
                   "../a.png", "a.png ", " a.png", "a\n.png",
                   "x" * 256 + ".png", str(self.root / "a.png"))
        for name in invalid:
            with self.subTest(name=name):
                error = self.reject(frames_meta=[{"file": name}])
                self.assertIn("invalid PNG filename", error)

    def test_symlink_alias_cannot_count_as_second_capture(self):
        (self.root / "a.png").write_bytes(b"fixture")
        try:
            (self.root / "alias.png").symlink_to(self.root / "a.png")
        except (OSError, NotImplementedError):
            self.skipTest("symlinks unavailable on this platform")
        error = self.reject(frames_meta=[
            {"file": "a.png", "timeMs": 0, "frame": 0},
            {"file": "alias.png", "timeMs": 1000, "frame": 30},
        ])
        self.assertIn("duplicate resolved", error)

    def test_directory_scan_has_bound_even_for_unrelated_entries(self):
        fake = repeat(SimpleNamespace(name="unrelated.txt"), 100_001)
        with patch.object(gemini_video.os, "scandir",
                          return_value=nullcontext(fake)):
            error = self.reject(max_images=2)
        self.assertIn("100000 entry scan limit", error)

    def test_unsampled_empty_frame_without_metadata_fails(self):
        for i in range(5):
            (self.root / f"{i}.png").write_bytes(b"" if i == 1 else b"fixture")
        error = self.reject(max_images=3)
        self.assertIn("empty", error)

    def test_unsampled_oversized_frame_without_metadata_fails(self):
        for i in range(5):
            (self.root / f"{i}.png").write_bytes(b"fixture")
        real_getsize = gemini_video.os.path.getsize
        def size(path):
            if Path(path).name == "1.png":
                return gemini_video.MAX_FRAME_IMAGE_BYTES + 1
            return real_getsize(path)
        with patch.object(gemini_video.os.path, "getsize", side_effect=size):
            error = self.reject(max_images=3)
        self.assertIn("upload size limit", error)


if __name__ == "__main__":
    unittest.main()
