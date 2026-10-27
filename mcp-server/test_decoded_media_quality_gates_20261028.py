"""Regression tests for opt-in export codec and signal-quality gates (#61)."""
import json
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import decoded_media_qa as qa


class ExportQualityGateTests(unittest.TestCase):
    def setUp(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.video = Path(directory.name) / "export.mp4"
        self.video.write_bytes(b"test export media " * 8)

    def inspect(self, *, video_codec="h264", audio_codec="aac", include_audio=True,
                video_log="", audio_log="", video_frames=None, **options):
        streams = [{"codec_type": "video", "start_time": "0",
                    "duration": "4"}]
        if video_frames is not None:
            streams[0]["nb_frames"] = video_frames
        if video_codec is not None:
            streams[0]["codec_name"] = video_codec
        if include_audio:
            audio = {"codec_type": "audio", "start_time": "0",
                     "duration": "4", "channels": 2}
            if audio_codec is not None:
                audio["codec_name"] = audio_codec
            streams.append(audio)
        metadata = {"format": {"duration": "4", "start_time": "0"},
                    "streams": streams}

        def fake_run(command, timeout):
            if command[0] == "ffprobe":
                return subprocess.CompletedProcess(command, 0, json.dumps(metadata), "")
            is_audio = "0:a:0" in command
            return subprocess.CompletedProcess(
                command, 0, "", audio_log if is_audio else video_log)

        with patch.object(qa, "run", side_effect=fake_run):
            return qa.inspect(str(self.video), sample_seconds=4,
                              require_audio=include_audio, **options)

    def test_expected_video_codec_accepts_matching_track(self):
        report = self.inspect(expected_video_codec="h264")
        self.assertTrue(report["passed"], report)
        self.assertEqual(report["video_details"][0]["codec"], "h264")

    def test_expected_video_codec_rejects_mismatch_and_missing(self):
        for codec, code in (("hevc", "EXPORT_VIDEO_CODEC_MISMATCH"),
                            (None, "EXPORT_VIDEO_CODEC_UNKNOWN")):
            with self.subTest(codec=codec):
                report = self.inspect(video_codec=codec, expected_video_codec="h264")
                self.assertFalse(report["passed"], report)
                self.assertIn(code, {i["code"] for i in report["issues"]})

    def test_expected_audio_codec_rejects_mismatch_and_missing(self):
        for codec, code in (("opus", "EXPORT_AUDIO_CODEC_MISMATCH"),
                            (None, "EXPORT_AUDIO_CODEC_UNKNOWN")):
            with self.subTest(codec=codec):
                report = self.inspect(audio_codec=codec, expected_audio_codec="aac")
                self.assertFalse(report["passed"], report)
                self.assertIn(code, {i["code"] for i in report["issues"]})

    def test_expected_audio_codec_requires_an_audio_track(self):
        report = self.inspect(include_audio=False, expected_audio_codec="aac")
        self.assertFalse(report["passed"], report)
        self.assertIn("EXPORT_AUDIO_STREAM_MISSING",
                      {i["code"] for i in report["issues"]})

    def test_invalid_codec_metadata_fails_closed(self):
        for codec in ("H264", "h264 extra", "😀"):
            with self.subTest(codec=codec):
                report = self.inspect(video_codec=codec)
                self.assertEqual(report.get("error_code"), "EXPORT_PROBE_FAILED")

    def test_black_duration_gate_rejects_observed_black_frames(self):
        report = self.inspect(include_audio=False, max_black_seconds=1.0,
                              video_log="black_start: 0 black_end: 2 black_duration: 2")
        self.assertFalse(report["passed"], report)
        self.assertEqual(report["verification_scope"], "incomplete")
        self.assertIn("EXPORT_BLACK_DURATION_EXCEEDED",
                      {i["code"] for i in report["issues"]})

    def test_silence_duration_gate_rejects_observed_long_silence(self):
        report = self.inspect(max_silence_seconds=1.0,
                              audio_log="silence_start: 0 silence_end: 2 silence_duration: 2")
        self.assertFalse(report["passed"], report)
        self.assertIn("EXPORT_SILENCE_DURATION_EXCEEDED",
                      {i["code"] for i in report["issues"]})

    def test_freeze_duration_gate_rejects_observed_freeze(self):
        report = self.inspect(include_audio=False, max_freeze_seconds=2.0,
                              video_log="freeze_start: 0 freeze_end: 3 freeze_duration: 3")
        self.assertFalse(report["passed"], report)
        self.assertIn("EXPORT_FREEZE_DURATION_EXCEEDED",
                      {i["code"] for i in report["issues"]})

    def test_quality_gate_preserves_warning_only_default(self):
        report = self.inspect(include_audio=False,
                              video_log="black_start: 0 black_end: 2 black_duration: 2")
        self.assertTrue(report["passed"], report)
        self.assertEqual(report["signal_scope"], "sampled_only")

    def test_strict_audio_metrics_rejects_unmeasured_audio(self):
        report = self.inspect(require_audio_peak_metrics=True)
        self.assertFalse(report["passed"], report)
        self.assertIn("EXPORT_SAMPLE_AUDIO_METRICS_UNAVAILABLE",
                      {i["code"] for i in report["issues"]})

    def test_strict_audio_metrics_accepts_valid_silence_measurement(self):
        report = self.inspect(require_audio_peak_metrics=True,
                              audio_log="Overall\nPeak level dB: -inf")
        self.assertTrue(report["passed"], report)

    def test_invalid_quality_gate_arguments_rejected_before_ffprobe(self):
        for options in ({"expected_video_codec": "H264"},
                        {"expected_audio_codec": "aac", "expected_audio_streams": 0,
                         "require_audio": False},
                        {"max_black_seconds": True},
                        {"max_silence_seconds": -1},
                        {"max_freeze_seconds": float("inf")},
                        {"require_audio_peak_metrics": 1}):
            with self.subTest(options=options), self.assertRaises(ValueError):
                # Pass require_audio only once, through the fixture helper.
                require_audio = options.pop("require_audio", None)
                if require_audio is None:
                    self.inspect(**options)
                else:
                    self.inspect(include_audio=require_audio, **options)

    def test_expected_video_frame_count_checks_all_tracks(self):
        report = self.inspect(video_frames="120", expected_video_frames=120)
        self.assertTrue(report["passed"], report)
        for frames, code in (("119", "EXPORT_VIDEO_FRAME_COUNT_MISMATCH"),
                             (None, "EXPORT_VIDEO_FRAME_COUNT_UNKNOWN")):
            with self.subTest(frames=frames):
                report = self.inspect(video_frames=frames, expected_video_frames=120)
                self.assertFalse(report["passed"], report)
                self.assertIn(code, {i["code"] for i in report["issues"]})

    def test_audio_peak_ceiling_blocks_loud_samples(self):
        report = self.inspect(max_audio_peak_dbfs=-1.0,
                              audio_log="Overall\nPeak level dB: -0.05")
        self.assertFalse(report["passed"], report)
        self.assertIn("EXPORT_SAMPLE_AUDIO_PEAK_TOO_HIGH",
                      {i["code"] for i in report["issues"]})

    def test_audio_peak_floor_blocks_quiet_and_silent_samples(self):
        for log in ("Overall\nPeak level dB: -45",
                    "Overall\nPeak level dB: -inf"):
            with self.subTest(log=log):
                report = self.inspect(min_audio_peak_dbfs=-20, audio_log=log)
                self.assertFalse(report["passed"], report)
                self.assertIn("EXPORT_SAMPLE_AUDIO_PEAK_TOO_LOW",
                              {i["code"] for i in report["issues"]})

    def test_peak_gate_rejects_missing_measurement(self):
        report = self.inspect(max_audio_peak_dbfs=-1.0)
        self.assertFalse(report["passed"], report)
        self.assertIn("EXPORT_SAMPLE_AUDIO_PEAK_UNVERIFIABLE",
                      {i["code"] for i in report["issues"]})

    def test_stereo_balance_gate_rejects_unbalanced_channels(self):
        log = ("Channel: 1\nPeak level dB: -5\n"
               "Channel: 2\nPeak level dB: -30\n"
               "Overall\nPeak level dB: -5")
        report = self.inspect(max_stereo_imbalance_db=12, audio_log=log)
        self.assertFalse(report["passed"], report)
        self.assertIn("EXPORT_SAMPLE_STEREO_IMBALANCE_EXCEEDED",
                      {i["code"] for i in report["issues"]})

    def test_stereo_balance_gate_requires_channel_measurements(self):
        report = self.inspect(max_stereo_imbalance_db=12,
                              audio_log="Overall\nPeak level dB: -5")
        self.assertFalse(report["passed"], report)
        self.assertIn("EXPORT_SAMPLE_STEREO_METRICS_UNAVAILABLE",
                      {i["code"] for i in report["issues"]})

    def test_additional_gate_arguments_are_strict(self):
        for options in ({"expected_video_frames": True},
                        {"expected_video_frames": 0},
                        {"min_audio_peak_dbfs": float("nan")},
                        {"max_audio_peak_dbfs": -130},
                        {"max_stereo_imbalance_db": -1},
                        {"min_audio_peak_dbfs": -1,
                         "max_audio_peak_dbfs": -10}):
            with self.subTest(options=options), self.assertRaises(ValueError):
                self.inspect(**options)


if __name__ == "__main__":
    unittest.main()
