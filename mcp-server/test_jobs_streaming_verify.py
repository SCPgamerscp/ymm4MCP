"""Regression tests for bounded export verification and invalid path controls."""
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import jobs


class BoundedExportVerificationTests(unittest.TestCase):
    def test_large_export_uses_bounded_header_read(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "large.mp4"
            with path.open("wb") as stream:
                stream.write(b"\x00\x00\x00\x20ftypisom")
                stream.truncate(16 * 1024 * 1024)
            with patch.object(Path, "read_bytes", side_effect=AssertionError("unbounded read")):
                result = jobs.verify_media_file(str(path), "mp4")
            self.assertTrue(result["ok"])
            self.assertEqual(result["bytes"], 16 * 1024 * 1024)
            self.assertEqual(result["brand"], "isom")

    def test_disappearing_export_is_reported_as_missing(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "race.wav"
            path.write_bytes(b"RIFF" + b"\x00" * 32)
            with patch.object(Path, "open", side_effect=FileNotFoundError("removed")):
                result = jobs.verify_media_file(str(path), "wav")
            self.assertEqual(result["error_code"], "EXPORT_FILE_MISSING")

    def test_unreadable_export_is_reported_as_verification_failure(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "unreadable.mp4"
            path.write_bytes(b"\x00" * 64)
            with patch.object(Path, "open", side_effect=PermissionError("denied")):
                result = jobs.verify_media_file(str(path), "mp4")
            self.assertEqual(result["error_code"], "EXPORT_VERIFY_FAILED")

    def test_control_characters_rejected_before_export_or_project_open(self):
        for char in ("\x00", "\r", "\n", "\t", "\x1f", "\x7f"):
            with self.subTest(char=repr(char)):
                self.assertFalse(jobs.is_absolute_media_path("C:/Videos/sub" + char + "/out.mp4"))
                with self.assertRaises(ValueError):
                    jobs.validate_export_request({"path": "C:/Videos/sub" + char + "/out.mp4"})
                with self.assertRaises(ValueError):
                    jobs.validate_project_path({"path": "C:/Projects/sub" + char + "/edit.ymmp"})


if __name__ == "__main__":
    unittest.main()
