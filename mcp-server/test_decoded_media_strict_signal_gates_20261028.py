"""Regression tests for strict signal thresholds and audio QA (#61)."""
import json
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import decoded_media_qa as qa


class StrictSignalGateTests(unittest.TestCase):
    def setUp(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.path = Path(directory.name) / "render.mp4"
        self.path.write_bytes(b"test export media " * 8)

    def inspect_with_commands(self, *, include_audio=True, video_log="", audio_log="",
                              full_signal_scan=False, **options):
        streams = [{"codec_type": "video", "start_time": "0",
                    "duration": "4"}]
        if include_audio:
            streams.append({"codec_type": "audio", "start_time": "0",
                            "duration": "4", "channels": 2})
        metadata = {"format": {"duration": "4", "start_time": "0"},
                    "streams": streams}
        commands = []

        def fake_run(command, timeout):
            commands.append(command)
            if command[0] == "ffprobe":
                return subprocess.CompletedProcess(command, 0, json.dumps(metadata), "")
            is_audio = "-af" in command
            return subprocess.CompletedProcess(command, 0, "",
                                               audio_log if is_audio else video_log)

        with patch.object(qa, "run", side_effect=fake_run):
            report = qa.inspect(str(self.path), sample_seconds=4,
                                require_audio=include_audio,
                                full_decode=full_signal_scan,
                                full_signal_scan=full_signal_scan, **options)
        return report, commands

    @staticmethod
    def filters(commands, flag):
        return [command[command.index(flag) + 1]
                for command in commands if flag in command]

    def test_strict_black_gate_reduces_sample_filter_minimum(self):
        report, commands = self.inspect_with_commands(
            include_audio=False, max_black_seconds=0.25,
            video_log="black_start: 0 black_end: 0.5 black_duration: 0.5")
        self.assertFalse(report["passed"], report)
        self.assertIn("EXPORT_BLACK_DURATION_EXCEEDED",
                      {item["code"] for item in report["issues"]})
        self.assertIn("blackdetect=d=0.250000", self.filters(commands, "-vf")[0])

    def test_strict_freeze_gate_reduces_sample_filter_minimum(self):
        report, commands = self.inspect_with_commands(
            include_audio=False, max_freeze_seconds=0.75,
            video_log="freeze_start: 0 freeze_end: 1 freeze_duration: 1")
        self.assertFalse(report["passed"], report)
        self.assertIn("EXPORT_FREEZE_DURATION_EXCEEDED",
                      {item["code"] for item in report["issues"]})
        self.assertIn("freezedetect=n=-60dB:d=0.750000",
                      self.filters(commands, "-vf")[0])

    def test_strict_silence_gate_reduces_sample_filter_minimum(self):
        report, commands = self.inspect_with_commands(
            max_silence_seconds=0.25,
            audio_log="silence_start: 0 silence_end: 0.5 silence_duration: 0.5")
        self.assertFalse(report["passed"], report)
        self.assertIn("EXPORT_SILENCE_DURATION_EXCEEDED",
                      {item["code"] for item in report["issues"]})
        self.assertIn("silencedetect=noise=-50dB:d=0.250000",
                      self.filters(commands, "-af")[0])

    def test_zero_tolerance_uses_bounded_one_millisecond_minimum(self):
        report, commands = self.inspect_with_commands(
            max_black_seconds=0, max_freeze_seconds=0, max_silence_seconds=0)
        self.assertTrue(report["passed"], report)
        self.assertIn("blackdetect=d=0.001000", self.filters(commands, "-vf")[0])
        self.assertIn("freezedetect=n=-60dB:d=0.001000",
                      self.filters(commands, "-vf")[0])
        self.assertIn("silencedetect=noise=-50dB:d=0.001000",
                      self.filters(commands, "-af")[0])

    def test_default_thresholds_preserve_legacy_detection_minima(self):
        report, commands = self.inspect_with_commands()
        self.assertTrue(report["passed"], report)
        self.assertIn("blackdetect=d=1.000000", self.filters(commands, "-vf")[0])
        self.assertIn("freezedetect=n=-60dB:d=2.000000",
                      self.filters(commands, "-vf")[0])
        self.assertIn("silencedetect=noise=-50dB:d=2.000000",
                      self.filters(commands, "-af")[0])

    def test_full_signal_scan_uses_same_strict_minima(self):
        report, commands = self.inspect_with_commands(
            full_signal_scan=True, max_black_seconds=0.4,
            max_freeze_seconds=0.6, max_silence_seconds=0.8)
        self.assertTrue(report["passed"], report)
        self.assertEqual(report["signal_scope"], "full")
        video_filters = self.filters(commands, "-vf")
        audio_filters = self.filters(commands, "-af")
        self.assertEqual(len(video_filters), 2)
        self.assertEqual(len(audio_filters), 2)
        self.assertTrue(all("blackdetect=d=0.400000" in f
                            and "freezedetect=n=-60dB:d=0.600000" in f
                            for f in video_filters))
        self.assertTrue(all("silencedetect=noise=-50dB:d=0.800000" in f
                            for f in audio_filters))

    def test_silence_quality_gate_requires_audio_when_optional(self):
        report, _ = self.inspect_with_commands(
            include_audio=False, max_silence_seconds=1)
        self.assertFalse(report["passed"], report)
        self.assertIn("EXPORT_AUDIO_STREAM_MISSING",
                      {item["code"] for item in report["issues"]})

    def test_silence_gate_rejects_explicit_zero_audio_streams(self):
        with self.assertRaises(ValueError):
            self.inspect_with_commands(
                include_audio=False, expected_audio_streams=0,
                max_silence_seconds=1)

    def test_strict_audio_peak_requires_all_declared_channel_metrics(self):
        partial = ("Channel: 1\nPeak level dB: -8\n"
                   "Overall\nPeak level dB: -8")
        report, _ = self.inspect_with_commands(
            require_audio_peak_metrics=True, audio_log=partial)
        self.assertFalse(report["passed"], report)
        self.assertIn("EXPORT_SAMPLE_AUDIO_METRICS_INCOMPLETE",
                      {item["code"] for item in report["issues"]})

    def test_strict_audio_peak_accepts_complete_channel_metrics(self):
        complete = ("Channel: 1\nPeak level dB: -8\n"
                    "Channel: 2\nPeak level dB: -9\n"
                    "Overall\nPeak level dB: -8")
        report, _ = self.inspect_with_commands(
            require_audio_peak_metrics=True, audio_log=complete)
        self.assertTrue(report["passed"], report)


if __name__ == "__main__":
    unittest.main()
