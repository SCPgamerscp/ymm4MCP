"""Unprefixed FFmpeg astats diagnostics must not bypass export QA (#61)."""
import json
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import decoded_media_qa as qa


class BareAstatsDiagnosticsTests(unittest.TestCase):
    def setUp(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.video = Path(directory.name) / "sample.mp4"
        self.video.write_bytes(b"fake media bytes" * 4)

    def inspect_audio(self, log):
        metadata = json.dumps({
            "format": {"duration": "1.5"},
            "streams": [
                {"codec_type": "video"},
                {"codec_type": "audio", "channels": 2},
            ],
        })

        def fake_run(command, timeout):
            if command[0] == "ffprobe":
                return subprocess.CompletedProcess(command, 0, metadata, "")
            return subprocess.CompletedProcess(
                command, 0, "", log if "0:a:0" in command else "")

        with patch.object(qa, "run", side_effect=fake_run):
            return qa.inspect(str(self.video))

    def test_unprefixed_malformed_peak_labels_fail_closed(self):
        for line in ("Peak level dB -10", "Peak level dB = -10",
                     "Peak level dBFS: -10", "Peak level dB : -10"):
            with self.subTest(line=line):
                report = self.inspect_audio(line + "\n")
                self.assertFalse(report["passed"], report)
                self.assertIn("EXPORT_SAMPLE_AUDIO_PEAK_INVALID",
                              [issue["code"] for issue in report["issues"]])

    def test_unprefixed_malformed_channel_header_fails_closed(self):
        report = self.inspect_audio(
            "Channel: 1\nPeak level dB: -20\n"
            "Channel 2\nPeak level dB: -1\n")
        self.assertFalse(report["passed"], report)
        self.assertIn("EXPORT_SAMPLE_AUDIO_CHANNEL_HEADER_INVALID",
                      [issue["code"] for issue in report["issues"]])
        self.assertEqual(qa._audio_channel_peaks(
            "Channel: 1\nPeak level dB: -20\n"
            "Channel 2\nPeak level dB: -1\n"), {1: -20.0})

    def test_unprefixed_overall_silence_is_reported(self):
        log = "Overall\nPeak level dB: -inf\n"
        self.assertTrue(qa._audio_overall_silent(log))
        report = self.inspect_audio(log)
        self.assertTrue(report["passed"], report)
        self.assertIn("EXPORT_SAMPLE_AUDIO_ALL_CHANNELS_SILENT",
                      [issue["code"] for issue in report["issues"]])

    def test_malformed_channel_header_ends_overall_section(self):
        log = ("Overall\nPeak level dB: -inf\n"
               "Channel 2\nPeak level dB: -inf\n")
        self.assertFalse(qa._audio_overall_silent(log))
        report = self.inspect_audio(log)
        self.assertFalse(report["passed"], report)
        self.assertNotIn("EXPORT_SAMPLE_AUDIO_ALL_CHANNELS_SILENT",
                         [issue["code"] for issue in report["issues"]])

    def test_unprefixed_overall_does_not_override_audible_peak(self):
        log = ("Overall\nPeak level dB: -6\n"
               "Overall\nPeak level dB: -inf\n")
        self.assertFalse(qa._audio_overall_silent(log))
        report = self.inspect_audio(log)
        self.assertTrue(report["passed"], report)
        self.assertEqual(report["audio_peak_db"], -6.0)
        self.assertNotIn("EXPORT_SAMPLE_AUDIO_ALL_CHANNELS_SILENT",
                         [issue["code"] for issue in report["issues"]])

    def test_unrelated_prefixed_filter_lines_are_not_trusted(self):
        log = ("[Parsed_other_0 @ 0x123] Peak level dBFS: -1\n"
               "[Parsed_other_0 @ 0x123] Channel 0\n"
               "[Parsed_other_0 @ 0x123] Overall\n"
               "[Parsed_other_0 @ 0x123] Peak level dB: -inf\n")
        report = self.inspect_audio(log)
        self.assertTrue(report["passed"], report)
        self.assertIsNone(report["audio_peak_db"])
        self.assertNotIn("EXPORT_SAMPLE_AUDIO_ALL_CHANNELS_SILENT",
                         [issue["code"] for issue in report["issues"]])


if __name__ == "__main__":
    unittest.main()
