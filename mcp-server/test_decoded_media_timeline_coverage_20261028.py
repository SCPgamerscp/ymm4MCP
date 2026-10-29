"""Regressions for export timeline boundaries and honest signal coverage (#61)."""
import json
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import decoded_media_qa as qa


class ExportTimelineCoverageTests(unittest.TestCase):
    def setUp(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.video = Path(directory.name) / "export.mp4"
        self.video.write_bytes(b"fake media for isolated ffprobe/ffmpeg mocks" * 2)

    def inspect(self, metadata, *, full_signal_scan=False, require_audio=False,
                fail_audio_at=None):
        def fake_run(command, timeout):
            if command[0] == "ffprobe":
                return subprocess.CompletedProcess(command, 0, json.dumps(metadata), "")
            if (fail_audio_at is not None and "-ss" in command
                    and "0:a:0" in command
                    and command[command.index("-ss") + 1] == fail_audio_at):
                return subprocess.CompletedProcess(command, 1, "", "decode failed")
            return subprocess.CompletedProcess(command, 0, "", "")

        with patch.object(qa, "run", side_effect=fake_run):
            return qa.inspect(str(self.video), sample_seconds=2,
                              require_audio=require_audio,
                              full_decode=full_signal_scan,
                              full_signal_scan=full_signal_scan)

    def test_track_starting_at_or_after_container_end_is_rejected(self):
        for start in ("4", "4.5"):
            with self.subTest(start=start):
                metadata = {
                    "format": {"duration": "4", "start_time": "0"},
                    "streams": [{"codec_type": "video", "start_time": start,
                                 "duration": "1"}],
                }
                report = self.inspect(metadata)
                self.assertFalse(report["success"], report)
                self.assertEqual(report["error_code"], "EXPORT_PROBE_FAILED")

    def test_track_ending_before_container_start_is_rejected(self):
        metadata = {
            "format": {"duration": "4", "start_time": "0"},
            "streams": [{"codec_type": "video", "start_time": "-3",
                         "duration": "2"}],
        }
        report = self.inspect(metadata)
        self.assertFalse(report["success"], report)
        self.assertEqual(report["error_code"], "EXPORT_PROBE_FAILED")

    def test_preroll_that_overlaps_container_is_still_valid(self):
        metadata = {
            "format": {"duration": "4", "start_time": "0"},
            "streams": [{"codec_type": "video", "start_time": "-2",
                         "duration": "6"}],
        }
        report = self.inspect(metadata)
        self.assertTrue(report["passed"], report)
        self.assertEqual(report["sampled_stream_count"], 1)

    def test_full_signal_scan_requires_known_start_and_duration(self):
        for missing in ("start_time", "duration"):
            with self.subTest(missing=missing):
                stream = {"codec_type": "video", "start_time": "0",
                          "duration": "4"}
                del stream[missing]
                metadata = {
                    "format": {"duration": "4", "start_time": "0"},
                    "streams": [stream],
                }
                report = self.inspect(metadata, full_signal_scan=True)
                self.assertFalse(report["passed"], report)
                self.assertEqual(report["signal_scope"], "sampled_only")
                self.assertIn("EXPORT_FULL_SIGNAL_TIMING_UNVERIFIABLE",
                              {issue["code"] for issue in report["issues"]})

    def test_full_signal_scan_with_complete_timing_is_full(self):
        metadata = {
            "format": {"duration": "4", "start_time": "0"},
            "streams": [{"codec_type": "video", "start_time": "0",
                         "duration": "4"}],
        }
        report = self.inspect(metadata, full_signal_scan=True)
        self.assertTrue(report["passed"], report)
        self.assertEqual(report["signal_scope"], "full")

    def test_sampled_stream_counts_are_per_track_and_per_success(self):
        metadata = {
            "format": {"duration": "4", "start_time": "0"},
            "streams": [{"codec_type": "video", "start_time": "0",
                         "duration": "4"},
                        {"codec_type": "audio", "start_time": "0",
                         "duration": "4"}],
        }
        report = self.inspect(metadata, require_audio=True,
                              fail_audio_at="2.000000")
        self.assertFalse(report["passed"], report)
        self.assertEqual(report["sampled_stream_count"], 2)
        self.assertEqual(report["sampled_streams"], [
            {"stream_type": "video", "stream_ordinal": 0, "sample_count": 2},
            {"stream_type": "audio", "stream_ordinal": 0, "sample_count": 1},
        ])
        self.assertIn("EXPORT_DECODE_FAILED",
                      {issue["code"] for issue in report["issues"]})


if __name__ == "__main__":
    unittest.main()
