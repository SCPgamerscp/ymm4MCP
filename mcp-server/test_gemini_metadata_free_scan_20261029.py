"""Safety regression tests for metadata-free YMM4 preview PNG discovery (#61)."""
import os
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock, patch

import gemini_video


class FakeScan:
    def __init__(self, *names):
        self.names = names

    def __enter__(self):
        return iter(SimpleNamespace(name=name) for name in self.names)

    def __exit__(self, *_):
        return False


class MetadataFreeCaptureTests(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)
        self.available = patch.object(
            gemini_video, "is_available", return_value=(True, "ok"))
        self.available.start()
        self.addCleanup(self.available.stop)

    def reject(self, **kwargs):
        with patch.object(gemini_video, "_client",
                          side_effect=AssertionError("Gemini must not be called")):
            result = gemini_video.analyze_frames_dir(str(self.root), **kwargs)
        self.assertFalse(result["success"], result)
        return result["error"]

    def test_reserved_windows_device_and_ads_names_are_invalid(self):
        for name in ("CON.png", "NUL.png", "COM1.png", "LPT².png",
                     "CON.anything.png", "clip.png:secret.png",
                     "frame?.png", "frame\n.png", " frame.png"):
            with self.subTest(name=name):
                self.assertFalse(gemini_video._valid_capture_png_name(name))
                error = self.reject(frames_meta=[{"file": name}])
                self.assertIn("invalid PNG filename", error)
        for name in ("0001.png", "シーン01.PNG", "scene 1.png"):
            with self.subTest(name=name):
                self.assertTrue(gemini_video._valid_capture_png_name(name))

    def test_directory_scan_rejects_invalid_windows_filename(self):
        with patch.object(gemini_video.os, "scandir",
                          return_value=FakeScan("CON.png")):
            error = self.reject()
        self.assertIn("invalid filename", error)

    def test_directory_scan_rejects_case_only_duplicates(self):
        with patch.object(gemini_video.os, "scandir",
                          return_value=FakeScan("A.png", "a.png")):
            error = self.reject()
        self.assertIn("case-insensitive duplicate", error)

    def test_directory_scan_rejects_hardlinked_captures(self):
        (self.root / "a.png").write_bytes(b"synthetic-a")
        try:
            os.link(self.root / "a.png", self.root / "b.png")
        except (OSError, NotImplementedError):
            self.skipTest("hardlinks unavailable")
        error = self.reject()
        self.assertIn("duplicate physical", error)

    def test_directory_scan_rejects_symlink_aliases(self):
        (self.root / "a.png").write_bytes(b"synthetic-a")
        try:
            (self.root / "b.png").symlink_to(self.root / "a.png")
        except (OSError, NotImplementedError):
            self.skipTest("symlinks unavailable")
        error = self.reject()
        self.assertIn("duplicate resolved", error)

    def test_truncated_or_appended_wav_is_rejected_before_model_call(self):
        (self.root / "a.png").write_bytes(b"synthetic-a")
        audio = self.root / "sound.wav"
        client = MagicMock()
        part = SimpleNamespace(
            from_bytes=lambda **kwargs: kwargs["data"],
            from_text=lambda **kwargs: kwargs["text"])
        for content in (
            b"RIFF" + (100).to_bytes(4, "little") + b"WAVE",
            b"RIFF" + (4).to_bytes(4, "little") + b"WAVEextra",
            b"RIFF" + (3).to_bytes(4, "little") + b"WAVE",
        ):
            with self.subTest(content=content):
                audio.write_bytes(content)
                with patch.object(gemini_video, "_client", return_value=client), \
                     patch.object(gemini_video, "genai_types",
                                  SimpleNamespace(Part=part)):
                    result = gemini_video.analyze_frames_dir(
                        str(self.root), audio_path=str(audio))
                self.assertFalse(result["success"], result)
                self.assertIn("RIFF/WAVE size", result["error"])
        client.models.generate_content.assert_not_called()

    def test_distinct_regular_frames_are_accepted(self):
        (self.root / "a.png").write_bytes(b"synthetic-a")
        (self.root / "b.png").write_bytes(b"synthetic-b")
        client = MagicMock()
        client.models.generate_content.return_value = SimpleNamespace(
            text='{"summary":"ok","events":[]}')
        part = SimpleNamespace(
            from_bytes=lambda **kwargs: kwargs["data"],
            from_text=lambda **kwargs: kwargs["text"])
        with patch.object(gemini_video, "_client", return_value=client), \
             patch.object(gemini_video, "genai_types",
                          SimpleNamespace(Part=part)):
            result = gemini_video.analyze_frames_dir(str(self.root))
        self.assertTrue(result["success"], result)
        self.assertEqual(result["analyzed_images"], 2)
        client.models.generate_content.assert_called_once()


if __name__ == "__main__":
    unittest.main()
