"""Regressions for FFmpeg signal chronology, durations and track timing (#61)."""
import json
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import decoded_media_qa as qa


class SignalTimingIntegrityTests(unittest.TestCase):
    def setUp(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.video = Path(directory.name) / "clip.mp4"
        self.video.write_bytes(b"synthetic video bytes" * 4)

    def inspect(self, *, video_sample="", audio_sample="", video_full="",
                audio_full="", metadata=None, full_signal_scan=False):
        if metadata is None:
            metadata = {
                "format": {"duration": "3", "start_time": "10"},
                "streams": [
                    {"codec_type": "video", "index": 0,
                     "start_time": "10", "duration": "3"},
                    {"codec_type": "audio", "index": 1,
                     "start_time": "10.5", "duration": "2"},
                ],
            }

        def fake_run(command, timeout):
            if command[0] == "ffprobe":
                return subprocess.CompletedProcess(
                    command, 0, json.dumps(metadata), "")
            if "-ss" in command:
                log = (video_sample if "0:v:0" in command else audio_sample)
            elif "-vf" in command or "-af" in command:
                log = video_full if "-vf" in command else audio_full
            else:
                log = ""
            return subprocess.CompletedProcess(command, 0, "", log)

        with patch.object(qa, "run", side_effect=fake_run):
            return qa.inspect(str(self.video), sample_seconds=3,
                              full_decode=full_signal_scan,
                              full_signal_scan=full_signal_scan)

    def test_backward_black_end_is_an_error_not_an_eof_blackout(self):
        result = self.inspect(video_sample=(
            "black_start: 1.5 black_end: 0.5\n"))
        self.assertFalse(result["passed"], result)
        self.assertIn("EXPORT_SAMPLE_SIGNAL_ORDER_INVALID",
                      {issue["code"] for issue in result["issues"]})

    def test_backward_silence_and_freeze_end_are_errors(self):
        for kind, kwargs in (
                ("silence", {"audio_sample":
                             "silence_start: 2 silence_end: 1\n"}),
                ("freeze", {"video_sample":
                            "freeze_start: 2 freeze_end: 1\n"})):
            with self.subTest(kind=kind):
                result = self.inspect(**kwargs)
                self.assertFalse(result["passed"], result)
                self.assertIn("EXPORT_SAMPLE_SIGNAL_ORDER_INVALID",
                              {issue["code"] for issue in result["issues"]})

    def test_duplicate_starts_and_unpaired_ends_remain_compatible(self):
        result = self.inspect(video_sample=(
            "black_end: 0.2\nblack_start: 0.5\n"
            "black_start: 0.8\nblack_end: 2.5\n"))
        self.assertTrue(result["passed"], result)
        self.assertEqual(len(result["black_frames"]), 1)
        self.assertEqual(result["black_frames"][0]["duration_seconds"], 2.0)

    def test_malformed_signal_duration_fails_for_each_filter(self):
        for kind, kwargs in (
                ("black", {"video_sample":
                           "black_start: 0 black_end: 2 black_duration: -2\n"}),
                ("silence", {"audio_sample":
                             "silence_end: 2 | silence_duration: NaN\n"}),
                ("freeze", {"video_sample":
                            "freeze_end: 2 | freeze_duration: 1e999\n"}),
                ("suffix", {"video_sample":
                            "black_duration: 2,broken\n"})):
            with self.subTest(kind=kind):
                result = self.inspect(**kwargs)
                self.assertFalse(result["passed"], result)
                self.assertIn("EXPORT_SAMPLE_SIGNAL_DURATION_INVALID",
                              {issue["code"] for issue in result["issues"]})

    def test_valid_durations_and_unrelated_filters_do_not_fail(self):
        result = self.inspect(
            video_sample=(
                "[Parsed_blackdetect_0 @ 0x123] black_start: 0 black_end: 2 black_duration: 2\n"
                "[Parsed_freezedetect_0 @ 0x123] freeze_start: 0 freeze_end: 2 freeze_duration: 2\n"
                "[Parsed_other_0 @ 0x123] black_duration: -999\n"),
            audio_sample=(
                "[Parsed_silencedetect_0 @ 0x123] silence_start: 0\n"
                "[Parsed_silencedetect_0 @ 0x123] silence_end: 2 | silence_duration: 2\n"))
        self.assertTrue(result["passed"], result)
        self.assertNotIn("EXPORT_SAMPLE_SIGNAL_DURATION_INVALID",
                         {issue["code"] for issue in result["issues"]})

    def test_full_signal_scan_rejects_backward_event(self):
        result = self.inspect(full_signal_scan=True, video_full=(
            "[Parsed_blackdetect_0 @ 0x123] black_start: 2 black_end: 1\n"))
        self.assertFalse(result["passed"], result)
        self.assertEqual(result["signal_scope"], "sampled_only")
        self.assertIn("EXPORT_FULL_SIGNAL_ORDER_INVALID",
                      {issue["code"] for issue in result["issues"]})

    def test_full_signal_scan_rejects_invalid_duration(self):
        result = self.inspect(full_signal_scan=True, audio_full=(
            "[Parsed_silencedetect_0 @ 0x123] silence_duration: -1\n"))
        self.assertFalse(result["passed"], result)
        self.assertIn("EXPORT_FULL_SIGNAL_DURATION_INVALID",
                      {issue["code"] for issue in result["issues"]})

    def test_contradictory_reported_signal_durations_fail_closed(self):
        for kind, kwargs in (
                ("black", {"video_sample":
                           "black_start: 0 black_end: 2 black_duration: 1\\n"}),
                ("silence", {"audio_sample":
                             "silence_start: 0\\nsilence_end: 2 | silence_duration: 1\\n"}),
                ("freeze", {"video_sample":
                            "freeze_start: 0\\nfreeze_end: 2 | freeze_duration: 1\\n"})):
            with self.subTest(kind=kind):
                result = self.inspect(**kwargs)
                self.assertFalse(result["passed"], result)
                self.assertIn("EXPORT_SAMPLE_SIGNAL_DURATION_MISMATCH",
                              {issue["code"] for issue in result["issues"]})

    def test_full_signal_duration_disagreement_fails_closed(self):
        result = self.inspect(full_signal_scan=True, audio_full=(
            "silence_start: 0\\nsilence_end: 2 | silence_duration: 0.1\\n"))
        self.assertFalse(result["passed"], result)
        self.assertIn("EXPORT_FULL_SIGNAL_DURATION_MISMATCH",
                      {issue["code"] for issue in result["issues"]})

    def test_small_duration_rounding_differences_are_tolerated(self):
        result = self.inspect(video_sample=(
            "black_start: 0.001 black_end: 2.000 black_duration: 2.000\\n"))
        self.assertTrue(result["passed"], result)

    def test_track_details_include_container_relative_start_and_end(self):
        result = self.inspect()
        self.assertTrue(result["passed"], result)
        video, audio = result["video_details"][0], result["audio_details"][0]
        self.assertEqual((video["start_seconds"], video["duration_seconds"],
                          video["end_seconds"]), (0.0, 3.0, 3.0))
        self.assertEqual((audio["start_seconds"], audio["duration_seconds"],
                          audio["end_seconds"]), (0.5, 2.0, 2.5))

    def test_unknown_track_timing_remains_unknown(self):
        metadata = {
            "format": {"duration": "3"},
            "streams": [{"codec_type": "video"},
                        {"codec_type": "audio", "start_time": "2"}],
        }
        result = self.inspect(metadata=metadata)
        self.assertTrue(result["passed"], result)
        video, audio = result["video_details"][0], result["audio_details"][0]
        self.assertIsNone(video["start_seconds"])
        self.assertIsNone(video["duration_seconds"])
        self.assertIsNone(video["end_seconds"])
        self.assertEqual(audio["start_seconds"], 2.0)
        self.assertIsNone(audio["end_seconds"])


if __name__ == "__main__":
    unittest.main()
