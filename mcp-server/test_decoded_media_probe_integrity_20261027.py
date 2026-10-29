"""Regression tests for contradictory ffprobe size and time-base metadata (#61)."""
import json
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import decoded_media_qa as qa


class ProbeIntegrityTests(unittest.TestCase):
    def setUp(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.video = Path(directory.name) / "clip.mp4"
        self.video.write_bytes(b"sample media content" * 4)

    def inspect(self, *, size=None, time_base=None):
        fmt = {"duration": "3.0", "start_time": "0"}
        if size is not None:
            fmt["size"] = size
        stream = {"index": 0, "codec_type": "video",
                  "start_time": "0", "duration": "3.0"}
        if time_base is not None:
            stream["time_base"] = time_base
        metadata = {"format": fmt, "streams": [stream]}

        def fake_run(command, timeout):
            if command[0] == "ffprobe":
                return subprocess.CompletedProcess(
                    command, 0, json.dumps(metadata), "")
            return subprocess.CompletedProcess(command, 0, "", "")

        with patch.object(qa, "run", side_effect=fake_run):
            return qa.inspect(str(self.video), require_audio=False)

    def test_matching_size_and_valid_time_base_pass(self):
        result = self.inspect(size=str(self.video.stat().st_size),
                              time_base="1/90000")
        self.assertTrue(result["passed"], result)

    def test_missing_optional_size_and_time_base_still_pass(self):
        result = self.inspect()
        self.assertTrue(result["passed"], result)

    def test_stale_or_forged_size_is_rejected(self):
        result = self.inspect(size=str(self.video.stat().st_size + 1))
        self.assertEqual(result["error_code"], "EXPORT_PROBE_FAILED")

    def test_malformed_size_is_rejected(self):
        for value in (True, 42, "-42", "42 ", "4e1", "1" * 30):
            with self.subTest(value=value):
                result = self.inspect(size=value)
                self.assertEqual(result["error_code"], "EXPORT_PROBE_FAILED")

    def test_zero_or_malformed_time_base_is_rejected(self):
        for value in ("0/1000", "1/0", "0/0", "1/-90000",
                      "1/90000 ", "N/A", "bad", True, "1/" + "9" * 30):
            with self.subTest(value=value):
                result = self.inspect(time_base=value)
                self.assertEqual(result["error_code"], "EXPORT_PROBE_FAILED")

    def test_valid_nonstandard_time_base_is_supported(self):
        for value in ("1001/30000", "1/1000000000"):
            with self.subTest(value=value):
                result = self.inspect(time_base=value)
                self.assertTrue(result["passed"], result)


if __name__ == "__main__":
    unittest.main()
