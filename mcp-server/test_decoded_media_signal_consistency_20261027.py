"""Regression tests for short-track sample coverage and contradictory astats (#61)."""
import json
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import decoded_media_qa as qa


class SignalConsistencyTests(unittest.TestCase):
    def setUp(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.video = Path(directory.name) / "clip.mp4"
        self.video.write_bytes(b"fake media bytes" * 4)

    def inspect(self, metadata, audio_log=""):
        calls = []

        def fake_run(command, timeout):
            calls.append(command)
            if command[0] == "ffprobe":
                return subprocess.CompletedProcess(
                    command, 0, json.dumps(metadata), "")
            log = audio_log if "-map" in command and command[
                command.index("-map") + 1] == "0:a:0" else ""
            return subprocess.CompletedProcess(command, 0, "", log)

        with patch.object(qa, "run", side_effect=fake_run):
            return qa.inspect(str(self.video)), calls

    def test_edge_only_overlap_adds_targeted_track_sample(self):
        metadata = {
            "format": {"duration": "30", "start_time": "0"},
            "streams": [
                {"codec_type": "video", "start_time": "0", "duration": "30"},
                {"codec_type": "audio", "start_time": "7.99",
                 "duration": "0.51", "channels": 2},
            ],
        }
        report, calls = self.inspect(metadata)
        self.assertTrue(report["passed"], report)
        starts = [item["start_seconds"] for item in report["samples"]]
        self.assertEqual(len(starts), 4, starts)
        self.assertIn(0.0, starts)
        self.assertIn(11.0, starts)
        self.assertIn(22.0, starts)
        self.assertTrue(any(0 < start < 11 for start in starts), starts)
        self.assertTrue(any(
            command[0] == "ffmpeg" and "0:a:0" in command and
            0 < float(command[command.index("-ss") + 1]) < 11
            for command in calls), calls)

    def test_meaningful_existing_overlap_does_not_add_redundant_sample(self):
        metadata = {
            "format": {"duration": "30", "start_time": "0"},
            "streams": [
                {"codec_type": "video", "start_time": "0", "duration": "30"},
                {"codec_type": "audio", "start_time": "7.7",
                 "duration": "0.8"},
            ],
        }
        report, _ = self.inspect(metadata)
        self.assertTrue(report["passed"], report)
        self.assertEqual(
            [item["start_seconds"] for item in report["samples"]],
            [0.0, 11.0, 22.0])

    def test_audible_channel_overrides_silent_overall(self):
        metadata = {"format": {"duration": "1.5"}, "streams": [
            {"codec_type": "video"}, {"codec_type": "audio", "channels": 2}]}
        log = ("[Parsed_astats_0 @ 0x123] Overall\n"
               "[Parsed_astats_0 @ 0x123] Peak level dB: -inf\n"
               "[Parsed_astats_0 @ 0x123] Channel: 1\n"
               "[Parsed_astats_0 @ 0x123] Peak level dB: -4\n"
               "[Parsed_astats_0 @ 0x123] Channel: 2\n"
               "[Parsed_astats_0 @ 0x123] Peak level dB: -inf\n")
        report, _ = self.inspect(metadata, log)
        self.assertTrue(report["passed"], report)
        self.assertNotIn("EXPORT_SAMPLE_AUDIO_ALL_CHANNELS_SILENT",
                         [item["code"] for item in report["issues"]])
        self.assertEqual(report["audio_peaks"][0]["peak_dbfs"], -4.0)

    def test_audible_overall_overrides_silent_channel_inventory(self):
        metadata = {"format": {"duration": "1.5"}, "streams": [
            {"codec_type": "video"}, {"codec_type": "audio", "channels": 2}]}
        log = ("[Parsed_astats_0 @ 0x123] Channel: 1\n"
               "[Parsed_astats_0 @ 0x123] Peak level dB: -inf\n"
               "[Parsed_astats_0 @ 0x123] Channel: 2\n"
               "[Parsed_astats_0 @ 0x123] Peak level dB: -inf\n"
               "[Parsed_astats_0 @ 0x123] Overall\n"
               "[Parsed_astats_0 @ 0x123] Peak level dB: -3\n")
        report, _ = self.inspect(metadata, log)
        self.assertTrue(report["passed"], report)
        self.assertNotIn("EXPORT_SAMPLE_AUDIO_ALL_CHANNELS_SILENT",
                         [item["code"] for item in report["issues"]])
        self.assertEqual(report["audio_peaks"][0]["peak_dbfs"], -3.0)


    def test_cross_filter_bare_log_does_not_spoof_black_events(self):
        log = ("silence_start: 0 black_start: 1\n"
               "silence_end: 2 black_end: 3\n")
        self.assertEqual(qa._signal_ranges(log, "black_start", "black_end", 4), [])
        self.assertEqual(qa._signal_ranges(log, "silence_start", "silence_end", 4),
                         [(0.0, 2.0)])

    def test_audio_peak_db_summary_uses_loudest_sample(self):
        metadata = {"format": {"duration": "30"}, "streams": [
            {"codec_type": "video"}, {"codec_type": "audio", "channels": 2}]}
        calls = []
        def fake_run(command, timeout):
            calls.append(command)
            if command[0] == "ffprobe":
                return subprocess.CompletedProcess(command, 0, json.dumps(metadata), "")
            if "0:a:0" in command:
                start = float(command[command.index("-ss") + 1])
                peak = -8 if start == 0 else -3 if start == 11 else -12
                return subprocess.CompletedProcess(
                    command, 0, "",
                    f"[Parsed_astats_0 @ 0x123] Peak level dB: {peak}\n")
            return subprocess.CompletedProcess(command, 0, "", "")
        with patch.object(qa, "run", side_effect=fake_run):
            report = qa.inspect(str(self.video))
        self.assertTrue(report["passed"], report)
        self.assertEqual(report["audio_peak_db"], -3.0)
        self.assertEqual(len(report["audio_peaks"]), 3)

    def test_audio_peak_db_unknown_without_finite_measurement(self):
        metadata = {"format": {"duration": "1.5"}, "streams": [
            {"codec_type": "video"}, {"codec_type": "audio"}]}
        report, _ = self.inspect(metadata)
        self.assertTrue(report["passed"], report)
        self.assertIsNone(report["audio_peak_db"])


    def test_stream_metadata_exposes_exact_fps_and_real_indices(self):
        metadata = {"format": {"duration": "1.5"}, "streams": [
            {"codec_type": "video", "index": 0,
             "disposition": {"attached_pic": 1}},
            {"codec_type": "video", "index": 2,
             "avg_frame_rate": "60000/2002", "nb_frames": "45"},
            {"codec_type": "audio", "index": 5, "sample_rate": "48000"}]}
        report, calls = self.inspect(metadata)
        self.assertTrue(report["passed"], report)
        self.assertEqual(report["video_details"][0]["stream_index"], 2)
        self.assertEqual(report["video_details"][0]["fps_ratio"], "30000/1001")
        self.assertEqual(report["video_details"][0]["reported_frames"], 45)
        self.assertEqual(report["audio_details"][0]["stream_index"], 5)
        self.assertTrue(any("0:2" in command for command in calls
                            if command[0] == "ffmpeg"))

    def test_unknown_fps_ratio_remains_null(self):
        metadata = {"format": {"duration": "1.5"}, "streams": [
            {"codec_type": "video", "avg_frame_rate": "0/0"},
            {"codec_type": "audio"}]}
        report, _ = self.inspect(metadata)
        self.assertTrue(report["passed"], report)
        self.assertIsNone(report["video_details"][0]["fps_ratio"])


if __name__ == "__main__":
    unittest.main()
