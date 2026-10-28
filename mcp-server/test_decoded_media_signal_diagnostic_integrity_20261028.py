"""Regression coverage for malformed signal diagnostics and strict astats (#61)."""
import json
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import decoded_media_qa as qa


class SignalDiagnosticIntegrityTests(unittest.TestCase):
    def setUp(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.path = Path(directory.name) / "render.mp4"
        self.path.write_bytes(b"test export media " * 8)

    def inspect_logs(self, video_log="", audio_log="", *, full_scan=False,
                     require_audio=False, audio_channels=2, **options):
        streams = [{"codec_type": "video", "start_time": "0", "duration": "4"}]
        if require_audio:
            audio = {"codec_type": "audio", "start_time": "0", "duration": "4"}
            if audio_channels is not None:
                audio["channels"] = audio_channels
            streams.append(audio)
        metadata = {"format": {"duration": "4", "start_time": "0"},
                    "streams": streams}

        def fake_run(command, timeout):
            if command[0] == "ffprobe":
                return subprocess.CompletedProcess(command, 0, json.dumps(metadata), "")
            return subprocess.CompletedProcess(
                command, 0, "", audio_log if "-af" in command else video_log)

        with patch.object(qa, "run", side_effect=fake_run):
            return qa.inspect(str(self.path), sample_seconds=4,
                              require_audio=require_audio,
                              full_decode=full_scan, full_signal_scan=full_scan,
                              **options)

    @staticmethod
    def codes(report):
        return {issue["code"] for issue in report["issues"]}

    def test_malformed_bare_black_start_is_not_silently_ignored(self):
        result = self.inspect_logs(video_log="black_start = 0.5")
        self.assertFalse(result["passed"], result)
        self.assertIn("EXPORT_SAMPLE_SIGNAL_LABEL_INVALID", self.codes(result))

    def test_malformed_prefixed_freeze_label_is_rejected(self):
        result = self.inspect_logs(
            video_log="[Parsed_freezedetect_1 @ 0x123] freeze_end 1.5")
        self.assertFalse(result["passed"], result)
        self.assertIn("EXPORT_SAMPLE_SIGNAL_LABEL_INVALID", self.codes(result))

    def test_malformed_bare_silence_duration_is_rejected(self):
        result = self.inspect_logs(require_audio=True,
                                   audio_log="silence_duration = 2")
        self.assertFalse(result["passed"], result)
        self.assertIn("EXPORT_SAMPLE_SIGNAL_LABEL_INVALID", self.codes(result))

    def test_unrelated_filter_cannot_inject_bad_signal_label(self):
        result = self.inspect_logs(
            video_log="[Parsed_other_0 @ 0x123] black_start = 0.5")
        self.assertTrue(result["passed"], result)

    def test_valid_signal_labels_remain_supported(self):
        result = self.inspect_logs(video_log=(
            "black_start: 0.5 black_end: 1.5 black_duration: 1"))
        self.assertTrue(result["passed"], result)
        self.assertNotIn("EXPORT_SAMPLE_SIGNAL_LABEL_INVALID", self.codes(result))

    def test_sample_rejects_black_timestamp_far_beyond_decode_window(self):
        result = self.inspect_logs(video_log="black_start: 20 black_end: 21")
        self.assertFalse(result["passed"], result)
        self.assertIn("EXPORT_SAMPLE_SIGNAL_TIMESTAMP_INVALID", self.codes(result))

    def test_sample_rejects_silence_timestamp_far_beyond_decode_window(self):
        result = self.inspect_logs(require_audio=True,
                                   audio_log="silence_start: 20 silence_end: 21")
        self.assertFalse(result["passed"], result)
        self.assertIn("EXPORT_SAMPLE_SIGNAL_TIMESTAMP_INVALID", self.codes(result))

    def test_sample_rejects_freeze_timestamp_far_beyond_decode_window(self):
        result = self.inspect_logs(video_log="freeze_start: 20 freeze_end: 21")
        self.assertFalse(result["passed"], result)
        self.assertIn("EXPORT_SAMPLE_SIGNAL_TIMESTAMP_INVALID", self.codes(result))

    def test_full_signal_scan_rejects_malformed_labels(self):
        result = self.inspect_logs(
            video_log="[Parsed_blackdetect_0 @ 0x123] black_start = 0.5",
            full_scan=True)
        self.assertFalse(result["passed"], result)
        self.assertIn("EXPORT_FULL_SIGNAL_LABEL_INVALID", self.codes(result))
        self.assertEqual(result["signal_scope"], "sampled_only")

    def test_full_signal_scan_rejects_out_of_bounds_timestamps(self):
        result = self.inspect_logs(
            video_log="freeze_start: 20 freeze_end: 21", full_scan=True)
        self.assertFalse(result["passed"], result)
        self.assertIn("EXPORT_FULL_SIGNAL_TIMESTAMP_INVALID", self.codes(result))

    def test_strict_peak_metrics_reject_channel_headers_without_peaks(self):
        result = self.inspect_logs(
            require_audio=True, require_audio_peak_metrics=True,
            audio_log=("Channel: 1\nChannel: 2\nOverall\nPeak level dB: -inf"))
        self.assertFalse(result["passed"], result)
        self.assertIn("EXPORT_SAMPLE_AUDIO_METRICS_INCOMPLETE",
                      self.codes(result))

    def test_strict_peak_metrics_reject_unknown_channel_count(self):
        result = self.inspect_logs(
            require_audio=True, audio_channels=None,
            require_audio_peak_metrics=True,
            audio_log="Overall\nPeak level dB: -12")
        self.assertFalse(result["passed"], result)
        self.assertIn("EXPORT_SAMPLE_AUDIO_CHANNEL_COUNT_UNKNOWN",
                      self.codes(result))

    def test_non_strict_peak_metrics_allow_unknown_channel_count(self):
        result = self.inspect_logs(
            require_audio=True, audio_channels=None,
            audio_log="Overall\nPeak level dB: -12")
        self.assertTrue(result["passed"], result)

    def test_strict_peak_metrics_accept_complete_silent_channels(self):
        result = self.inspect_logs(
            require_audio=True, require_audio_peak_metrics=True,
            audio_log=("Channel: 1\nPeak level dB: -inf\n"
                       "Channel: 2\nPeak level dB: -inf\n"
                       "Overall\nPeak level dB: -inf"))
        self.assertTrue(result["passed"], result)


if __name__ == "__main__":
    unittest.main()
