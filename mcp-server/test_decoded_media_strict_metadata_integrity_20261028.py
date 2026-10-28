"""Regression tests for fail-closed strict export quality gates (Issue #61)."""
import json
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import decoded_media_qa as qa


class StrictQAMetadataIntegrityTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.path = Path(self.directory.name) / "export.mp4"
        self.path.write_bytes(b"fake video data" * 8)

    def inspect(self, streams, *, audio_log="", **options):
        metadata = json.dumps({"format": {"duration": "4", "start_time": "0"},
                               "streams": streams})

        def fake_run(command, timeout):
            if command[0] == "ffprobe":
                return subprocess.CompletedProcess(command, 0, metadata, "")
            return subprocess.CompletedProcess(
                command, 0, "", audio_log if "-af" in command else "")

        with patch.object(qa, "run", side_effect=fake_run):
            return qa.inspect(str(self.path), sample_seconds=4,
                              require_audio=any(s["codec_type"] == "audio"
                                                for s in streams), **options)

    @staticmethod
    def codes(result):
        return {issue["code"] for issue in result["issues"]}

    def test_strict_video_gate_rejects_unknown_track_end(self):
        result = self.inspect(
            [{"codec_type": "video", "start_time": "0"}],
            max_black_seconds=0.2)
        self.assertFalse(result["passed"], result)
        self.assertIn("EXPORT_VIDEO_COVERAGE_UNVERIFIABLE", self.codes(result))

    def test_strict_audio_gate_rejects_unknown_track_origin(self):
        result = self.inspect(
            [{"codec_type": "video", "start_time": "0", "duration": "4"},
             {"codec_type": "audio", "duration": "4", "channels": 2}],
            max_silence_seconds=0.2)
        self.assertFalse(result["passed"], result)
        self.assertIn("EXPORT_AUDIO_COVERAGE_UNVERIFIABLE", self.codes(result))

    def test_default_gate_keeps_unknown_timing_backward_compatible(self):
        result = self.inspect([{"codec_type": "video"}])
        self.assertTrue(result["passed"], result)
        self.assertNotIn("EXPORT_VIDEO_COVERAGE_UNVERIFIABLE", self.codes(result))

    def test_strict_peak_rejects_overall_only_with_known_channels(self):
        result = self.inspect(
            [{"codec_type": "video"},
             {"codec_type": "audio", "channels": 2}],
            audio_log="Overall\nPeak level dB: -8\n",
            require_audio_peak_metrics=True)
        self.assertFalse(result["passed"], result)
        self.assertIn("EXPORT_SAMPLE_AUDIO_METRICS_INCOMPLETE", self.codes(result))

    def test_strict_peak_accepts_all_declared_channels(self):
        result = self.inspect(
            [{"codec_type": "video"},
             {"codec_type": "audio", "channels": 2}],
            audio_log=("Channel: 1\nPeak level dB: -8\n"
                       "Channel: 2\nPeak level dB: -9\n"
                       "Overall\nPeak level dB: -8\n"),
            require_audio_peak_metrics=True)
        self.assertTrue(result["passed"], result)

    def test_loudnorm_rejects_malformed_auxiliary_measurements(self):
        valid = ('{\n"input_i": "-20.5", "input_tp": "-3.0", '
                 '"input_lra": "1.5", "input_thresh": "-30.5", '
                 '"output_i": "-16.0"\n}')
        self.assertEqual(-20.5, qa._integrated_lufs_from_loudnorm(valid))
        for old, new in (('"input_tp": "-3.0"', '"input_tp": "nan"'),
                         ('"input_tp": "-3.0"', '"input_tp": "999"'),
                         ('"input_lra": "1.5"', '"input_lra": "-1"'),
                         ('"input_thresh": "-30.5"', '"input_thresh": "inf"'),
                         ('"output_i": "-16.0"', '"output_i": "NaN"'),
                         ('"output_i": "-16.0"', '"output_i": 0')):
            with self.subTest(replacement=new):
                with self.assertRaises(ValueError):
                    qa._integrated_lufs_from_loudnorm(valid.replace(old, new))


if __name__ == "__main__":
    unittest.main()
