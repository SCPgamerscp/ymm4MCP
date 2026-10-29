"""Fail-closed FFmpeg diagnostic regressions for completed export QA (#61)."""
import json
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import decoded_media_qa as qa


class MalformedFFmpegDiagnosticsTests(unittest.TestCase):
    def setUp(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.video = Path(directory.name) / "clip.mp4"
        self.video.write_bytes(b"synthetic video bytes" * 4)

    def check(self, *, video_log="", audio_log="", channels=2):
        audio = {"codec_type": "audio"}
        if channels is not None:
            audio["channels"] = channels
        metadata = json.dumps({
            "format": {"duration": "1.5"},
            "streams": [{"codec_type": "video"}, audio],
        })

        def fake_run(command, timeout):
            if command[0] == "ffprobe":
                return subprocess.CompletedProcess(command, 0, metadata, "")
            selected = command[command.index("-map") + 1]
            return subprocess.CompletedProcess(
                command, 0, "",
                video_log if selected == "0:v:0" else audio_log)

        with patch.object(qa, "run", side_effect=fake_run):
            return qa.inspect(str(self.video))

    def assert_error(self, report, code):
        self.assertFalse(report["passed"], report)
        self.assertIn(code, [issue["code"] for issue in report["issues"]], report)

    def test_malformed_astats_peaks_fail_closed(self):
        for raw in ("NaN", "inf", "+inf", "1e999", "garbled", "",
                    "-inf trailing"):
            with self.subTest(raw=raw):
                report = self.check(audio_log=(
                    "[Parsed_astats_0 @ 0x123] Overall\n"
                    "[Parsed_astats_0 @ 0x123] Peak level dB: " + raw + "\n"))
                self.assert_error(report, "EXPORT_SAMPLE_AUDIO_PEAK_INVALID")

    def test_finite_and_silent_astats_peaks_remain_valid(self):
        for raw in ("-18.25", "0", "-inf", "-Inf", "-INF"):
            with self.subTest(raw=raw):
                report = self.check(audio_log=(
                    "[Parsed_astats_0 @ 0x123] Overall\n"
                    "[Parsed_astats_0 @ 0x123] Peak level dB: " + raw + "\n"))
                self.assertTrue(report["passed"], report)
                self.assertNotIn("EXPORT_SAMPLE_AUDIO_PEAK_INVALID",
                                 [issue["code"] for issue in report["issues"]])

    def test_untrusted_peak_metadata_does_not_fail_qa(self):
        report = self.check(audio_log=(
            "title: Peak level dB: nan\n"
            "[other @ 0x123] Peak level dB: inf\n"))
        self.assertTrue(report["passed"], report)

    def test_invalid_black_and_freeze_events_fail_closed(self):
        for event, value in (
                ("black_start", "NaN"),
                ("black_end", "Infinity"),
                ("freeze_start", "1e999"),
                ("freeze_end", "broken"),
                ("black_start", "")):
            with self.subTest(event=event, value=value):
                filter_name = "blackdetect" if event.startswith("black") else "freezedetect"
                report = self.check(video_log=(
                    f"[Parsed_{filter_name}_0 @ 0x123] {event}: {value}\n"))
                self.assert_error(report, "EXPORT_SAMPLE_SIGNAL_TIMESTAMP_INVALID")

    def test_invalid_silence_events_fail_closed(self):
        for event, value in (("silence_start", "nan"),
                             ("silence_end", "1e999"),
                             ("silence_end", "")):
            with self.subTest(event=event, value=value):
                report = self.check(audio_log=(
                    f"[Parsed_silencedetect_0 @ 0x123] {event}: {value}\n"))
                self.assert_error(report, "EXPORT_SAMPLE_SIGNAL_TIMESTAMP_INVALID")

    def test_valid_signed_signal_timestamps_remain_supported(self):
        report = self.check(video_log=(
            "[Parsed_blackdetect_0 @ 0x123] black_start: -1.5 black_end: 1.0\n"
            "[Parsed_freezedetect_0 @ 0x123] freeze_start: 0.1 freeze_end: 1.0\n"))
        self.assertTrue(report["passed"], report)
        self.assertEqual(len(report["black_frames"]), 1, report)

    def test_unrelated_signal_metadata_is_not_trusted(self):
        report = self.check(video_log=(
            "title: black_start: NaN\n"
            "[Parsed_other_0 @ 0x123] freeze_start: 1e999\n"))
        self.assertTrue(report["passed"], report)

    def test_invalid_channel_headers_fail_even_without_declared_count(self):
        for header in ("Channel: 0", "Channel: 65",
                       "Channel: 999999999999", "Channel: nope"):
            with self.subTest(header=header):
                report = self.check(channels=None, audio_log=(
                    "[Parsed_astats_0 @ 0x123] " + header + "\n"))
                self.assert_error(report, "EXPORT_SAMPLE_AUDIO_CHANNEL_HEADER_INVALID")

    def test_valid_channel_headers_without_declared_count_remain_supported(self):
        report = self.check(channels=None, audio_log=(
            "[Parsed_astats_0 @ 0x123] Channel: 1\n"
            "[Parsed_astats_0 @ 0x123] Peak level dB: -10\n"))
        self.assertTrue(report["passed"], report)


if __name__ == "__main__":
    unittest.main()
