"""Regressions for strict container head/tail coverage (#61)."""
import json
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import decoded_media_qa as qa


class StrictContainerCoverageTests(unittest.TestCase):
    def setUp(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.path = Path(directory.name) / "export.mp4"
        self.path.write_bytes(b"fake media " * 16)

    def inspect_tracks(self, streams, **options):
        metadata = {"format": {"duration": "4", "start_time": "0"},
                    "streams": streams}

        def fake_run(command, timeout):
            if command[0] == "ffprobe":
                return subprocess.CompletedProcess(command, 0, json.dumps(metadata), "")
            return subprocess.CompletedProcess(command, 0, "", "")

        with patch.object(qa, "run", side_effect=fake_run):
            return qa.inspect(str(self.path), sample_seconds=4,
                              require_audio=any(
                                  stream["codec_type"] == "audio" for stream in streams),
                              **options)

    @staticmethod
    def codes(report):
        return {issue["code"] for issue in report["issues"]}

    def test_require_tracks_to_end_rejects_subsecond_tail_truncation(self):
        report = self.inspect_tracks([
            {"codec_type": "video", "start_time": "0", "duration": "3.7"}],
            require_tracks_to_end=True)
        self.assertFalse(report["passed"], report)
        self.assertIn("EXPORT_STREAM_ENDS_EARLY", self.codes(report))

    def test_require_tracks_to_end_allows_small_encoder_rounding(self):
        report = self.inspect_tracks([
            {"codec_type": "video", "start_time": "0", "duration": "3.97"}],
            require_tracks_to_end=True)
        self.assertTrue(report["passed"], report)

    def test_default_track_tail_remains_advisory_for_short_gap(self):
        report = self.inspect_tracks([
            {"codec_type": "video", "start_time": "0", "duration": "3.7"}])
        self.assertTrue(report["passed"], report)

    def test_strict_audio_gate_catches_unrendered_head(self):
        report = self.inspect_tracks([
            {"codec_type": "video", "start_time": "0", "duration": "4"},
            {"codec_type": "audio", "start_time": "0.4", "duration": "3.6"}],
            max_silence_seconds=0.2)
        self.assertFalse(report["passed"], report)
        self.assertIn("EXPORT_AUDIO_HEAD_GAP_EXCEEDED", self.codes(report))

    def test_strict_audio_gate_catches_unrendered_tail(self):
        report = self.inspect_tracks([
            {"codec_type": "video", "start_time": "0", "duration": "4"},
            {"codec_type": "audio", "start_time": "0", "duration": "3.6"}],
            max_silence_seconds=0.2)
        self.assertFalse(report["passed"], report)
        self.assertIn("EXPORT_AUDIO_TAIL_GAP_EXCEEDED", self.codes(report))

    def test_strict_video_gate_catches_unrendered_head(self):
        report = self.inspect_tracks([
            {"codec_type": "video", "start_time": "0.4", "duration": "3.6"}],
            max_black_seconds=0.2)
        self.assertFalse(report["passed"], report)
        self.assertIn("EXPORT_VIDEO_HEAD_GAP_EXCEEDED", self.codes(report))

    def test_strict_video_gate_catches_unrendered_tail(self):
        report = self.inspect_tracks([
            {"codec_type": "video", "start_time": "0", "duration": "3.6"}],
            max_black_seconds=0.2)
        self.assertFalse(report["passed"], report)
        self.assertIn("EXPORT_VIDEO_TAIL_GAP_EXCEEDED", self.codes(report))

    def test_staggered_tracks_cover_each_others_gaps(self):
        report = self.inspect_tracks([
            {"codec_type": "video", "start_time": "0.5", "duration": "3.5"},
            {"codec_type": "video", "start_time": "0", "duration": "2"}],
            max_black_seconds=0.1)
        self.assertTrue(report["passed"], report)
        self.assertNotIn("EXPORT_VIDEO_HEAD_GAP_EXCEEDED", self.codes(report))
        self.assertNotIn("EXPORT_VIDEO_TAIL_GAP_EXCEEDED", self.codes(report))

    def test_missing_track_timing_fails_closed_without_inventing_a_gap(self):
        report = self.inspect_tracks([
            {"codec_type": "video"}], max_black_seconds=0.1)
        self.assertFalse(report["passed"], report)
        self.assertIn("EXPORT_VIDEO_COVERAGE_UNVERIFIABLE", self.codes(report))
        self.assertNotIn("EXPORT_VIDEO_HEAD_GAP_EXCEEDED", self.codes(report))

    def test_small_gap_within_tolerance_is_allowed(self):
        report = self.inspect_tracks([
            {"codec_type": "video", "start_time": "0.12", "duration": "3.88"}],
            max_black_seconds=0.1)
        self.assertTrue(report["passed"], report)


    def test_video_gap_between_staggered_tracks_is_blocking(self):
        report = self.inspect_tracks([
            {"codec_type": "video", "start_time": "0", "duration": "1"},
            {"codec_type": "video", "start_time": "3", "duration": "1"}],
            max_black_seconds=0.2)
        self.assertFalse(report["passed"], report)
        self.assertIn("EXPORT_VIDEO_INTERNAL_GAP_EXCEEDED", self.codes(report))

    def test_audio_gap_between_staggered_tracks_is_blocking(self):
        report = self.inspect_tracks([
            {"codec_type": "video", "start_time": "0", "duration": "4"},
            {"codec_type": "audio", "start_time": "0", "duration": "1"},
            {"codec_type": "audio", "start_time": "3", "duration": "1"}],
            max_silence_seconds=0.2)
        self.assertFalse(report["passed"], report)
        self.assertIn("EXPORT_AUDIO_INTERNAL_GAP_EXCEEDED", self.codes(report))

    def test_overlapping_audio_tracks_have_no_internal_gap(self):
        report = self.inspect_tracks([
            {"codec_type": "video", "start_time": "0", "duration": "4"},
            {"codec_type": "audio", "start_time": "0", "duration": "3"},
            {"codec_type": "audio", "start_time": "2", "duration": "2"}],
            max_silence_seconds=0.2)
        self.assertTrue(report["passed"], report)
        self.assertNotIn("EXPORT_AUDIO_INTERNAL_GAP_EXCEEDED", self.codes(report))


if __name__ == "__main__":
    unittest.main()
