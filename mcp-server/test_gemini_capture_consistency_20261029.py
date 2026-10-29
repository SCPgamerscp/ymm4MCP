"""Regression coverage for contradictory YMM4 preview capture metadata (#61)."""
import os
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import gemini_video


class CaptureConsistencyTests(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.root = Path(tmp.name)
        available = patch.object(gemini_video, "is_available", return_value=(True, "ok"))
        available.start()
        self.addCleanup(available.stop)

    def frame(self, name):
        (self.root / name).write_bytes(b"synthetic-frame")

    def reject(self, metadata):
        with patch.object(gemini_video, "_client",
                          side_effect=AssertionError("Gemini must not be called")):
            result = gemini_video.analyze_frames_dir(
                str(self.root), frames_meta=metadata)
        self.assertFalse(result["success"], result)
        return result["error"]

    def test_case_only_aliases_rejected_on_every_platform(self):
        self.frame("A.png")
        self.frame("a.png")
        error = self.reject([
            {"file": "A.png", "timeMs": 0, "frame": 0},
            {"file": "a.png", "timeMs": 1000, "frame": 30},
        ])
        self.assertIn("case-insensitive duplicate", error)

    def test_hardlinks_do_not_count_as_two_captured_images(self):
        self.frame("a.png")
        try:
            os.link(self.root / "a.png", self.root / "b.png")
        except (OSError, NotImplementedError):
            self.skipTest("hard links unavailable on this filesystem")
        error = self.reject([
            {"file": "a.png", "timeMs": 0, "frame": 0},
            {"file": "b.png", "timeMs": 1000, "frame": 30},
        ])
        self.assertIn("duplicate physical", error)

    def test_conflicting_dimensions_rejected(self):
        self.frame("a.png")
        self.frame("b.png")
        error = self.reject([
            {"file": "a.png", "width": 1920, "height": 1080},
            {"file": "b.png", "width": 1280, "height": 720},
        ])
        self.assertIn("inconsistent capture dimensions", error)

    def test_capture_index_must_match_timestamp_order(self):
        self.frame("a.png")
        self.frame("b.png")
        error = self.reject([
            {"file": "a.png", "index": 1, "timeMs": 0, "frame": 0},
            {"file": "b.png", "index": 0, "timeMs": 1000, "frame": 30},
        ])
        self.assertIn("capture indices conflict", error)

    def test_capture_index_must_match_frame_order_without_times(self):
        self.frame("a.png")
        self.frame("b.png")
        error = self.reject([
            {"file": "a.png", "index": 1, "frame": 0},
            {"file": "b.png", "index": 0, "frame": 30},
        ])
        self.assertIn("capture indices conflict", error)

    def test_reject_oversized_or_malformed_instructions_before_upload(self):
        self.frame("a.png")
        for instruction in (12, ["bad"], "x" * 32_769, "hello" + chr(0) + "world",
                            "bad" + chr(0xD800) + "unicode"):
            with self.subTest(instruction_type=type(instruction).__name__):
                with patch.object(gemini_video, "_client",
                                  side_effect=AssertionError("Gemini must not be called")):
                    result = gemini_video.analyze_frames_dir(
                        str(self.root), event_instruction=instruction)
                self.assertFalse(result["success"], result)
                self.assertIn("event_instruction", result["error"])

    def test_png_preflight_stat_failure_returns_structured_error(self):
        self.frame("a.png")
        with patch.object(gemini_video.os.path, "getsize",
                          side_effect=OSError("stat denied")):
            error = self.reject([{"file": "a.png"}])
        self.assertIn("Cannot inspect PNG frame", error)

    def test_valid_shuffled_capture_still_passes(self):
        self.frame("a.png")
        self.frame("b.png")

        class Part:
            @staticmethod
            def from_bytes(*, data, mime_type):
                return data

            @staticmethod
            def from_text(*, text):
                return text

        client = SimpleNamespace(models=SimpleNamespace(
            generate_content=lambda **kwargs: SimpleNamespace(
                text='{"summary":"ok","events":[]}')))
        with patch.object(gemini_video, "_client", return_value=client), \
             patch.object(gemini_video, "genai_types", SimpleNamespace(Part=Part)):
            result = gemini_video.analyze_frames_dir(str(self.root), frames_meta=[
                {"file": "b.png", "index": 1, "timeMs": 1000, "frame": 30,
                 "width": 1920, "height": 1080},
                {"file": "a.png", "index": 0, "timeMs": 0, "frame": 0,
                 "width": 1920, "height": 1080},
            ])
        self.assertTrue(result["success"], result)
        self.assertEqual(result["analyzed_images"], 2)


if __name__ == "__main__":
    unittest.main()
