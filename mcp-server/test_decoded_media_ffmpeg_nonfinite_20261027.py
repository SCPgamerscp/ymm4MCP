"""Regression tests for non-finite FFmpeg signals and astats formatting."""
import json
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import decoded_media_qa as qa


class FFmpegSignalIntegrityTests(unittest.TestCase):
    def test_nonfinite_end_does_not_invent_signal_until_sample_eof(self):
        for token in ("NaN", "nan", "INF", "Infinity", "-inf", "1e999", "1e+", "invalid", "2junk"):
            with self.subTest(token=token):
                log = f"[blackdetect @ 0x1] black_start: 1 black_end: {token}"
                self.assertEqual(
                    qa._signal_ranges(log, "black_start", "black_end", 8), [])

    def test_nonfinite_end_does_not_invent_silence_or_freeze(self):
        for filter_name, start_name, end_name in (
            ("silencedetect", "silence_start", "silence_end"),
            ("freezedetect", "freeze_start", "freeze_end"),
        ):
            with self.subTest(filter_name=filter_name):
                log = f"[{filter_name} @ 0x1] {start_name}: 1 {end_name}: NaN"
                self.assertEqual(qa._signal_ranges(log, start_name, end_name, 8), [])

    def test_valid_signal_still_closes_and_unfinished_signal_reaches_eof(self):
        self.assertEqual(
            qa._signal_ranges("black_start: 1 black_end: 2",
                              "black_start", "black_end", 8), [(1.0, 2.0)])
        self.assertEqual(
            qa._signal_ranges("black_start: 1", "black_start", "black_end", 8),
            [(1.0, 8.0)])

    def test_overall_silence_allows_ffmpeg_spacing_and_case(self):
        for peak in ("Peak level dB: -inf", "Peak level dB:   -INF",
                     "Peak level dB:\t-Inf"):
            with self.subTest(peak=peak):
                log = "[Parsed_astats_0 @ 0x1] Overall\n" + (
                    "[Parsed_astats_0 @ 0x1] " + peak + "\n")
                self.assertTrue(qa._audio_overall_silent(log))

    def test_overall_finite_peak_is_not_silence(self):
        log = ("[Parsed_astats_0 @ 0x1] Overall\n"
               "[Parsed_astats_0 @ 0x1] Peak level dB:   -12\n")
        self.assertFalse(qa._audio_overall_silent(log))


class FFmpegChannelInventoryTests(unittest.TestCase):
    def _inspect_audio_log(self, audio_log):
        metadata = json.dumps({
            "format": {"duration": "1.5"},
            "streams": [{"codec_type": "video"},
                        {"codec_type": "audio", "channels": 2}],
        })
        with tempfile.TemporaryDirectory() as folder:
            video = Path(folder) / "clip.mp4"
            video.write_bytes(b"test media fixture" + b"x" * 32)

            def fake(command, timeout):
                if command[0] == "ffprobe":
                    return subprocess.CompletedProcess(command, 0, metadata, "")
                return subprocess.CompletedProcess(
                    command, 0, "",
                    audio_log if "-af" in command else "")

            with patch.object(qa, "run", side_effect=fake):
                return qa.inspect(str(video))

    def test_out_of_range_channel_header_is_reported_without_peak(self):
        report = self._inspect_audio_log(
            "[Parsed_astats_0 @ 0x1] Channel: 999\n")
        mismatches = [issue for issue in report["issues"]
                      if issue["code"] == "EXPORT_SAMPLE_AUDIO_CHANNEL_COUNT_MISMATCH"]
        self.assertEqual(len(mismatches), 1, report)
        self.assertTrue(mismatches[0]["invalid_channel_headers"])
        self.assertEqual(mismatches[0]["observed_channels"], [])

    def test_unexpected_channel_header_is_reported_without_peak(self):
        report = self._inspect_audio_log(
            "[Parsed_astats_0 @ 0x1] Channel: 3\n")
        mismatches = [issue for issue in report["issues"]
                      if issue["code"] == "EXPORT_SAMPLE_AUDIO_CHANNEL_COUNT_MISMATCH"]
        self.assertEqual(len(mismatches), 1, report)
        self.assertEqual(mismatches[0]["observed_channels"], [3])

    def test_valid_channel_headers_without_peaks_do_not_misreport_count(self):
        report = self._inspect_audio_log(
            "[Parsed_astats_0 @ 0x1] Channel: 1\n"
            "[Parsed_astats_0 @ 0x1] Channel: 2\n")
        self.assertFalse(any(
            issue["code"] == "EXPORT_SAMPLE_AUDIO_CHANNEL_COUNT_MISMATCH"
            for issue in report["issues"]), report)


if __name__ == "__main__":
    unittest.main()
