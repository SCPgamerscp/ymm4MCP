"""Regression tests for expected export duration and robust FFmpeg signal pairs (#61)."""
import json
import math
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import decoded_media_qa as qa
import server


class DurationExpectationTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.video = Path(temporary.name) / "export.mp4"
        self.video.write_bytes(b"fake exported media" * 3)

    def inspect(self, duration, **kwargs):
        metadata = json.dumps({
            "format": {"duration": str(duration)},
            "streams": [{"codec_type": "video"}],
        })

        def fake_run(command, timeout):
            return subprocess.CompletedProcess(
                command, 0, metadata if command[0] == "ffprobe" else "", "")

        with patch.object(qa, "run", side_effect=fake_run):
            return qa.inspect(str(self.video), require_audio=False, **kwargs)

    def test_truncated_export_is_an_error_with_measured_difference(self):
        report = self.inspect(28.0, expected_duration_seconds=30.0,
                              duration_tolerance_seconds=0.5)
        self.assertTrue(report["success"], report)
        self.assertFalse(report["passed"], report)
        issue = next(i for i in report["issues"]
                     if i["code"] == "EXPORT_DURATION_MISMATCH")
        self.assertEqual(issue["severity"], "error")
        self.assertEqual(issue["actual_duration_seconds"], 28.0)
        self.assertEqual(issue["expected_duration_seconds"], 30.0)
        self.assertEqual(issue["difference_seconds"], -2.0)
        self.assertEqual(issue["tolerance_seconds"], 0.5)
        self.assertEqual(report["verification_scope"], "incomplete")

    def test_overlong_export_is_an_error_too(self):
        report = self.inspect(33, expected_duration_seconds=30,
                              duration_tolerance_seconds=1)
        self.assertFalse(report["passed"], report)
        self.assertEqual(next(i["difference_seconds"] for i in report["issues"]
                              if i["code"] == "EXPORT_DURATION_MISMATCH"), 3.0)

    def test_within_tolerance_and_exact_boundary_are_accepted(self):
        for duration in (29.5, 30.0, 30.5):
            with self.subTest(duration=duration):
                report = self.inspect(duration, expected_duration_seconds=30,
                                      duration_tolerance_seconds=0.5)
                self.assertTrue(report["passed"], report)
                self.assertFalse(any(i["code"] == "EXPORT_DURATION_MISMATCH"
                                     for i in report["issues"]))

    def test_omitted_expectation_preserves_existing_behavior(self):
        report = self.inspect(1.5)
        self.assertTrue(report["passed"], report)
        self.assertEqual(report["issues"], [])

    def test_reject_invalid_expected_duration_and_tolerance(self):
        for value in (True, "30", 0, -1, math.inf, -math.inf, math.nan, 604801):
            with self.subTest(expected=value), self.assertRaises(ValueError):
                qa.inspect(str(self.video), require_audio=False,
                           expected_duration_seconds=value)
        for value in (True, "1", -0.1, math.inf, math.nan, 61):
            with self.subTest(tolerance=value), self.assertRaises(ValueError):
                qa.inspect(str(self.video), require_audio=False,
                           duration_tolerance_seconds=value)


class SignalPairTests(unittest.TestCase):
    def test_repeated_start_keeps_first_observed_boundary(self):
        log = "black_start: 1.0\\nblack_start: 2.0\\nblack_end: 4.0"
        self.assertEqual(qa._signal_ranges(log, "black_start", "black_end", 8),
                         [(1.0, 4.0)])

    def test_out_of_order_end_does_not_erase_active_start(self):
        log = "silence_start: 2\\nsilence_end: 1\\nsilence_end: 5"
        self.assertEqual(qa._signal_ranges(log, "silence_start", "silence_end", 8),
                         [(2.0, 5.0)])

    def test_duplicate_start_reaching_eof_closes_from_earliest_start(self):
        log = "freeze_start: 1\\nfreeze_start: 3"
        self.assertEqual(qa._signal_ranges(log, "freeze_start", "freeze_end", 6),
                         [(1.0, 6)])


class McpExpectationForwardingTests(unittest.IsolatedAsyncioTestCase):
    async def test_dispatch_forwards_duration_options(self):
        expected_report = {"success": True, "passed": True, "issues": []}
        with patch.object(server.decoded_media_qa, "inspect",
                          return_value=expected_report) as inspect:
            actual = await server.dispatch({
                "action": "get_info", "sub_action": "decoded_export_qa",
                "path": "/tmp/example.mp4",
                "expected_duration_seconds": 30,
                "duration_tolerance_seconds": 0.25,
            })
        self.assertEqual(actual, expected_report)
        inspect.assert_called_once_with(
            "/tmp/example.mp4", require_audio=True, sample_seconds=8,
            full_decode=False, expected_duration_seconds=30,
            duration_tolerance_seconds=0.25)


if __name__ == "__main__":
    unittest.main()
