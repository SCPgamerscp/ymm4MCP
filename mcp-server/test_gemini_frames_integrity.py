"""Regression tests for Gemini frame metadata and upload integrity."""
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch
import gemini_video


class GeminiFrameIntegrityTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        p = patch.object(gemini_video, "is_available", return_value=(True, "ok"))
        p.start()
        self.addCleanup(p.stop)

    def png(self, name, data=b"\x89PNG\r\n\x1a\nfixture"):
        (Path(self.tmp.name) / name).write_bytes(data)

    def reject_without_upload(self, meta):
        with patch.object(gemini_video, "_client", side_effect=AssertionError("unexpected upload")):
            result = gemini_video.analyze_frames_dir(self.tmp.name, frames_meta=meta)
        self.assertFalse(result["success"], result)
        return result["error"]

    def test_missing_or_invalid_filenames_are_not_skipped(self):
        self.png("a.png")
        for invalid in (None, "", 42, "other.jpg"):
            with self.subTest(invalid=invalid):
                error = self.reject_without_upload([
                    {"file": "a.png", "timeMs": 0, "frame": 0},
                    {"file": invalid, "timeMs": 1000, "frame": 30},
                ])
                self.assertIn("filename", error)

    def test_duplicate_filename_is_rejected(self):
        self.png("a.png")
        error = self.reject_without_upload([
            {"file": "a.png", "timeMs": 0, "frame": 0},
            {"file": "a.png", "timeMs": 1000, "frame": 30},
        ])
        self.assertIn("duplicate", error)

    def test_mixed_timestamps_are_not_hidden_by_valid_frame_indices(self):
        self.png("a.png")
        self.png("b.png")
        error = self.reject_without_upload([
            {"file": "a.png", "timeMs": 0, "frame": 0},
            {"file": "b.png", "timeMs": None, "frame": 30},
        ])
        self.assertIn("inconsistent", error)

    def test_conflicting_time_and_frame_order_is_rejected(self):
        self.png("a.png")
        self.png("b.png")
        error = self.reject_without_upload([
            {"file": "a.png", "timeMs": 1000, "frame": 0},
            {"file": "b.png", "timeMs": 0, "frame": 30},
        ])
        self.assertIn("conflicts", error)

    def test_duplicate_frame_index_is_rejected(self):
        self.png("a.png")
        self.png("b.png")
        error = self.reject_without_upload([
            {"file": "a.png", "timeMs": 0, "frame": 1},
            {"file": "b.png", "timeMs": 1000, "frame": 1},
        ])
        self.assertIn("conflicts", error)

    def test_directory_read_failure_is_structured(self):
        with patch.object(gemini_video.os, "listdir", side_effect=OSError("denied")):
            result = gemini_video.analyze_frames_dir(self.tmp.name)
        self.assertFalse(result["success"])
        self.assertIn("Cannot enumerate", result["error"])

    def test_natural_order_supports_very_long_numeric_filename(self):
        long_name = "9" * 5000 + ".png"
        self.assertEqual(sorted(["10.png", long_name, "2.png"],
                                key=gemini_video._frame_filename_key),
                         ["2.png", "10.png", long_name])

    def test_valid_shuffled_metadata_is_processed_in_order(self):
        self.png("a.png")
        self.png("b.png")
        labels = []
        class Part:
            @staticmethod
            def from_bytes(*, data, mime_type):
                return data
            @staticmethod
            def from_text(*, text):
                labels.append(text)
                return text
        client = SimpleNamespace(models=SimpleNamespace(
            generate_content=lambda **kwargs: SimpleNamespace(
                text='{"summary":"ok","events":[]}')))
        with patch.object(gemini_video, "_client", return_value=client), \
             patch.object(gemini_video, "genai_types", SimpleNamespace(Part=Part)):
            result = gemini_video.analyze_frames_dir(self.tmp.name, frames_meta=[
                {"file": "b.png", "timeMs": 1000, "frame": 30},
                {"file": "a.png", "timeMs": 0, "frame": 0},
            ])
        self.assertTrue(result["success"], result)
        self.assertIn("frame=0", labels[0])
        self.assertIn("frame=30", labels[1])


if __name__ == "__main__":
    unittest.main()
