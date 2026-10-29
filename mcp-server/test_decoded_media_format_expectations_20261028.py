"""Regression tests for expected decoded-export format and window deduplication."""
import json
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import decoded_media_qa as qa


class ExpectedDecodedFormatTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.video = Path(self.directory.name) / "clip.mp4"
        self.video.write_bytes(b"synthetic media" * 5)

    def check(self, streams, *, duration="30", **options):
        commands = []
        metadata = json.dumps({"format": {"duration": duration}, "streams": streams})

        def fake_run(command, timeout):
            commands.append(command)
            return subprocess.CompletedProcess(
                command, 0, metadata if command[0] == "ffprobe" else "", "")

        with patch.object(qa, "run", side_effect=fake_run):
            result = qa.inspect(str(self.video), **options)
        return result, commands

    def test_matching_video_and_audio_expectations_pass(self):
        streams = [
            {"codec_type": "video", "width": 1920, "height": 1080,
             "avg_frame_rate": "30000/1001"},
            {"codec_type": "audio", "sample_rate": "48000", "channels": 2},
        ]
        report, _ = self.check(
            streams, expected_width=1920, expected_height=1080,
            expected_fps=29.97, fps_tolerance=0.001,
            expected_audio_sample_rate=48000, expected_audio_channels=2)
        self.assertTrue(report["passed"], report)

    def test_video_mismatch_checks_all_tracks(self):
        streams = [
            {"codec_type": "video", "width": 1920, "height": 1080,
             "avg_frame_rate": "30/1"},
            {"codec_type": "video", "width": 1280, "height": 720,
             "avg_frame_rate": "25/1"},
        ]
        report, _ = self.check(
            streams, require_audio=False, expected_width=1920,
            expected_height=1080, expected_fps=30, fps_tolerance=0)
        self.assertFalse(report["passed"])
        mismatches = [i for i in report["issues"]
                      if i["code"] == "EXPORT_VIDEO_FORMAT_MISMATCH"]
        self.assertEqual({i["field"] for i in mismatches},
                         {"width", "height", "fps"})
        self.assertTrue(all(i["stream_ordinal"] == 1 for i in mismatches))

    def test_missing_video_metadata_is_not_treated_as_matching(self):
        report, _ = self.check(
            [{"codec_type": "video"}], require_audio=False,
            expected_width=1920, expected_height=1080, expected_fps=30)
        self.assertFalse(report["passed"])
        unknown = [i for i in report["issues"]
                   if i["code"] == "EXPORT_VIDEO_FORMAT_UNKNOWN"]
        self.assertEqual({i["field"] for i in unknown},
                         {"width", "height", "fps"})

    def test_audio_mismatch_checks_every_audio_track(self):
        streams = [
            {"codec_type": "video"},
            {"codec_type": "audio", "sample_rate": "48000", "channels": 2},
            {"codec_type": "audio", "sample_rate": "44100", "channels": 1},
        ]
        report, _ = self.check(
            streams, expected_audio_sample_rate=48000,
            expected_audio_channels=2)
        self.assertFalse(report["passed"])
        mismatches = [i for i in report["issues"]
                      if i["code"] == "EXPORT_AUDIO_FORMAT_MISMATCH"]
        self.assertEqual({i["field"] for i in mismatches},
                         {"sample_rate", "channels"})
        self.assertTrue(all(i["stream_ordinal"] == 1 for i in mismatches))

    def test_missing_audio_metadata_is_unverifiable(self):
        report, _ = self.check(
            [{"codec_type": "video"}, {"codec_type": "audio"}],
            expected_audio_sample_rate=48000, expected_audio_channels=2)
        self.assertFalse(report["passed"])
        self.assertEqual({i["field"] for i in report["issues"]
                          if i["code"] == "EXPORT_AUDIO_FORMAT_UNKNOWN"},
                         {"sample_rate", "channels"})

    def test_audio_expectation_requires_audio_even_if_optional(self):
        report, _ = self.check(
            [{"codec_type": "video"}], require_audio=False,
            expected_audio_channels=2)
        self.assertFalse(report["passed"])
        self.assertIn("EXPORT_AUDIO_STREAM_MISSING",
                      [i["code"] for i in report["issues"]])

    def test_unspecified_expectations_preserve_optional_audio(self):
        report, _ = self.check([{"codec_type": "video"}], require_audio=False)
        self.assertTrue(report["passed"], report)

    def test_expected_fps_uses_exact_ratio_not_rounded_display(self):
        streams = [{"codec_type": "video", "avg_frame_rate": "30000/1001"}]
        report, _ = self.check(streams, require_audio=False,
                               expected_fps=29.97003, fps_tolerance=0.00001)
        self.assertTrue(report["passed"], report)
        report, _ = self.check(streams, require_audio=False,
                               expected_fps=30, fps_tolerance=0.001)
        self.assertFalse(report["passed"])

    def test_expected_track_counts_reject_extra_video_and_audio(self):
        streams = [
            {"codec_type": "video"}, {"codec_type": "video"},
            {"codec_type": "audio"}, {"codec_type": "audio"},
        ]
        report, _ = self.check(
            streams, expected_video_streams=1, expected_audio_streams=1)
        self.assertFalse(report["passed"])
        self.assertEqual(
            {i["code"] for i in report["issues"]
             if i["code"].endswith("STREAM_COUNT_MISMATCH")},
            {"EXPORT_VIDEO_STREAM_COUNT_MISMATCH",
             "EXPORT_AUDIO_STREAM_COUNT_MISMATCH"})

    def test_expected_zero_audio_is_allowed_only_when_audio_optional(self):
        streams = [{"codec_type": "video"}]
        report, _ = self.check(
            streams, require_audio=False,
            expected_video_streams=1, expected_audio_streams=0)
        self.assertTrue(report["passed"], report)
        with self.assertRaises(ValueError):
            qa.inspect(str(self.video), expected_audio_streams=0)

    def test_expected_zero_audio_rejects_unexpected_audio_track(self):
        streams = [{"codec_type": "video"}, {"codec_type": "audio"}]
        report, _ = self.check(
            streams, require_audio=False, expected_audio_streams=0)
        self.assertFalse(report["passed"])
        self.assertIn("EXPORT_AUDIO_STREAM_COUNT_MISMATCH",
                      [i["code"] for i in report["issues"]])

    def test_invalid_expectations_fail_before_subprocess(self):
        invalid = (
            {"expected_width": True},
            {"expected_height": 0},
            {"expected_audio_sample_rate": "48000"},
            {"expected_audio_channels": 9},
            {"expected_fps": float("nan")},
            {"expected_fps": True},
            {"fps_tolerance": float("inf")},
            {"fps_tolerance": -1},
            {"expected_video_streams": 0},
            {"expected_video_streams": True},
            {"expected_audio_streams": -1},
            {"expected_audio_streams": 17},
        )
        for options in invalid:
            with self.subTest(options=options), self.assertRaises(ValueError):
                qa.inspect(str(self.video), **options)

    def test_identical_short_track_windows_are_not_decoded_twice(self):
        streams = [
            {"codec_type": "video", "start_time": "0", "duration": "30"},
            {"codec_type": "audio", "start_time": "7.99", "duration": "0.51"},
            {"codec_type": "audio", "start_time": "7.99", "duration": "0.51"},
        ]
        report, commands = self.check(streams)
        self.assertTrue(report["passed"], report)
        starts = [s["start_seconds"] for s in report["samples"]]
        self.assertEqual(len(starts), 4, starts)
        self.assertEqual(len(starts), len(set(starts)))
        # Video is decoded in four windows; each short audio track in two.
        self.assertEqual(len([c for c in commands if c[0] == "ffmpeg"]), 8)


if __name__ == "__main__":
    unittest.main()
