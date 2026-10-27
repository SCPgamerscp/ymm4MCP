"""Fail-closed QA regressions for malformed ffprobe format and impossible expectations (#61)."""
import json
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import decoded_media_qa as qa


class FailClosedExportInputsTests(unittest.TestCase):
    def setUp(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.video = Path(directory.name) / "clip.mp4"
        self.video.write_bytes(b"synthetic export" * 5)

    def check_probe(self, metadata):
        def fake_run(command, timeout):
            if command[0] == "ffprobe":
                return subprocess.CompletedProcess(
                    command, 0, json.dumps(metadata), "")
            return subprocess.CompletedProcess(command, 0, "", "")

        with patch.object(qa, "run", side_effect=fake_run):
            return qa.inspect(str(self.video), require_audio=False)

    def test_non_object_format_is_structured_probe_failure(self):
        # Previously `format.get(...)` raised AttributeError, escaping the
        # QA endpoint rather than returning a structured verification failure.
        for malformed in (None, [], [1], 123, "not a format", True):
            with self.subTest(format=malformed):
                report = self.check_probe({
                    "format": malformed,
                    "streams": [{"codec_type": "video"}],
                })
                self.assertFalse(report["passed"])
                self.assertEqual(report["error_code"], "EXPORT_PROBE_FAILED")

    def test_missing_format_is_structured_probe_failure(self):
        report = self.check_probe({
            "streams": [{"codec_type": "video"}],
        })
        self.assertFalse(report["passed"])
        self.assertEqual(report["error_code"], "EXPORT_PROBE_FAILED")

    def test_valid_format_object_remains_supported(self):
        report = self.check_probe({
            "format": {"duration": "2.0"},
            "streams": [{"codec_type": "video"}],
        })
        self.assertTrue(report["passed"], report)

    def test_extreme_integer_expectations_fail_with_value_error(self):
        # math.isfinite(10**400) raises OverflowError on CPython. These are
        # invalid caller options and must not become unhandled server errors.
        huge = 10 ** 400
        for options in (
            {"expected_duration_seconds": huge},
            {"duration_tolerance_seconds": huge},
            {"expected_fps": huge},
            {"fps_tolerance": huge},
        ):
            with self.subTest(options=list(options)), patch.object(
                    qa, "run") as run, self.assertRaises(ValueError):
                qa.inspect(str(self.video), **options)
                run.assert_not_called()

    def test_zero_audio_count_cannot_request_audio_format(self):
        for options in (
            {"expected_audio_sample_rate": 48000},
            {"expected_audio_channels": 2},
            {"expected_audio_sample_rate": 48000,
             "expected_audio_channels": 2},
        ):
            with self.subTest(options=options), patch.object(
                    qa, "run") as run, self.assertRaises(ValueError):
                qa.inspect(str(self.video), require_audio=False,
                           expected_audio_streams=0, **options)
                run.assert_not_called()

    def test_zero_audio_count_without_audio_format_is_valid(self):
        report = self.check_probe({
            "format": {"duration": "2.0"},
            "streams": [{"codec_type": "video"}],
        })
        self.assertTrue(report["passed"], report)


if __name__ == "__main__":
    unittest.main()
