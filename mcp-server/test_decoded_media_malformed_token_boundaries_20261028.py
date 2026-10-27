"""Malformed FFmpeg diagnostic token boundaries must fail closed (#61)."""
import json
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import decoded_media_qa as qa


class MalformedDiagnosticTokenTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.video = Path(temporary.name) / "clip.mp4"
        self.video.write_bytes(b"test media bytes" * 4)

    def inspect(self, *, video_log="", audio_log=""):
        metadata = json.dumps({
            "format": {"duration": "2.0"},
            "streams": [{"codec_type": "video"},
                        {"codec_type": "audio", "channels": 2}],
        })

        def fake_run(command, timeout):
            if command[0] == "ffprobe":
                return subprocess.CompletedProcess(command, 0, metadata, "")
            selected = command[command.index("-map") + 1]
            return subprocess.CompletedProcess(
                command, 0, "", video_log if selected == "0:v:0" else audio_log)

        with patch.object(qa, "run", side_effect=fake_run):
            return qa.inspect(str(self.video))

    def assert_issue(self, result, code):
        self.assertFalse(result["passed"], result)
        self.assertIn(code, [item["code"] for item in result["issues"]])

    def test_black_start_with_comma_suffix_fails_closed(self):
        report = self.inspect(video_log=(
            "[Parsed_blackdetect_0 @ 0x123] black_start: 0.5,broken black_end: 1.0\n"))
        self.assert_issue(report, "EXPORT_SAMPLE_SIGNAL_TIMESTAMP_INVALID")
        self.assertEqual(report["black_frames"], [])

    def test_silence_end_with_semicolon_suffix_fails_closed(self):
        report = self.inspect(audio_log=(
            "[Parsed_silencedetect_0 @ 0x123] silence_start: 0.1 "
            "silence_end: 1.0;bad\n"))
        self.assert_issue(report, "EXPORT_SAMPLE_SIGNAL_TIMESTAMP_INVALID")
        self.assertEqual(report["long_silences"], [])

    def test_freeze_start_with_comma_suffix_fails_closed(self):
        report = self.inspect(video_log=(
            "[Parsed_freezedetect_0 @ 0x123] freeze_start: 0.2,bad "
            "freeze_end: 1.2\n"))
        self.assert_issue(report, "EXPORT_SAMPLE_SIGNAL_TIMESTAMP_INVALID")
        self.assertEqual(report["frozen_frames"], [])

    def test_malformed_astats_peak_label_fails_closed(self):
        for line in ("Peak level dB -10", "Peak level dB = -10",
                     "Peak level dBFS: -10", "Peak level dB : -10"):
            with self.subTest(line=line):
                report = self.inspect(audio_log=(
                    "[Parsed_astats_0 @ 0x123] " + line + "\n"))
                self.assert_issue(report, "EXPORT_SAMPLE_AUDIO_PEAK_INVALID")

    def test_valid_filter_diagnostics_still_pass(self):
        report = self.inspect(
            video_log="[Parsed_blackdetect_0 @ 0x123] black_start: 0.2 black_end: 1.2\n",
            audio_log="[Parsed_astats_0 @ 0x123] Peak level dB: -12.5\n")
        self.assertTrue(report["passed"], report)
        self.assertEqual(len(report["black_frames"]), 1)
        self.assertEqual(report["audio_peak_db"], -12.5)

    def test_unrelated_filter_metadata_does_not_fail_qa(self):
        report = self.inspect(
            video_log="[Parsed_other_0 @ 0x123] black_start: 0.5,bad\n",
            audio_log="[Parsed_other_0 @ 0x123] Peak level dB -10\n")
        self.assertTrue(report["passed"], report)


if __name__ == "__main__":
    unittest.main()
