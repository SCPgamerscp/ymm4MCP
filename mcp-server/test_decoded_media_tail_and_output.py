"""Regression tests for decode QA tail coverage and invalid subprocess output."""
import json
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import decoded_media_qa as qa


class DecodedQATailAndOutputTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.video = Path(self.directory.name) / "clip.mp4"
        self.video.write_bytes(b"fake media bytes" * 4)

    def check_windows(self, duration, sample_seconds, expected_starts):
        calls = []
        metadata = json.dumps({"format": {"duration": str(duration)},
                               "streams": [{"codec_type": "video"},
                                           {"codec_type": "audio"}]})

        def fake(command, timeout):
            calls.append(command)
            return subprocess.CompletedProcess(command, 0,
                                               metadata if command[0] == "ffprobe" else "",
                                               "")

        with patch.object(qa, "run", side_effect=fake):
            result = qa.inspect(str(self.video), sample_seconds=sample_seconds)
        self.assertTrue(result["passed"], result)
        self.assertEqual([sample["start_seconds"] for sample in result["samples"]],
                         expected_starts)
        for start in expected_starts:
            for kind in ("0:v:0", "0:a:0"):
                self.assertTrue(any("-map" in command and
                                    command[command.index("-map") + 1] == kind and
                                    float(command[command.index("-ss") + 1]) == start
                                    for command in calls if command[0] == "ffmpeg"),
                                (kind, start, calls))
        return result

    def test_tail_is_checked_when_duration_is_between_one_and_two_windows(self):
        result = self.check_windows(12.5, 8, [0.0, 4.5])
        self.assertEqual(result["samples"][-1]["duration_seconds"], 8.0)
        self.assertEqual(result["samples"][-1]["start_seconds"] +
                         result["samples"][-1]["duration_seconds"], 12.5)

    def test_short_video_still_needs_only_one_window(self):
        self.check_windows(8.0, 8, [0.0])

    def test_exact_two_windows_checks_second_half(self):
        self.check_windows(16.0, 8, [0.0, 8.0])

    def test_long_video_still_checks_last_window_not_the_middle(self):
        self.check_windows(30.5, 8, [0.0, 22.5])

    def test_subprocess_output_uses_replacement_for_invalid_encoding(self):
        with patch.object(qa.subprocess, "run",
                          return_value=subprocess.CompletedProcess([], 0, "", "")) as runner:
            qa.run(["ffprobe", "-version"], 3)
        self.assertEqual(runner.call_args.kwargs["errors"], "replace")
        self.assertTrue(runner.call_args.kwargs["text"])

    def inspect_with_probe(self, metadata):
        def fake(command, timeout):
            return subprocess.CompletedProcess(
                command, 0, metadata if command[0] == "ffprobe" else "", "")
        with patch.object(qa, "run", side_effect=fake):
            return qa.inspect(str(self.video))

    def test_duplicate_ffprobe_json_keys_fail_closed(self):
        for metadata in (
            '{"format":{"duration":"12","duration":"20"},'
            '"streams":[{"codec_type":"video"}]}',
            '{"format":{"duration":"12"},"streams":[{"codec_type":"video"}],'
            '"streams":[{"codec_type":"audio"}]}',
            '{"format":{"duration":"12"},"streams":['
            '{"codec_type":"video","codec_type":"audio"}]}',
        ):
            with self.subTest(metadata=metadata):
                report = self.inspect_with_probe(metadata)
                self.assertEqual(report["error_code"], "EXPORT_PROBE_FAILED")

    def test_malformed_ffprobe_stream_entries_fail_closed(self):
        for streams in ([None], ["video"], [{"codec_type": 1}],
                        [{"codec_type": ""}], [{"codec_type": "video"}, 12]):
            with self.subTest(streams=streams):
                report = self.inspect_with_probe(json.dumps(
                    {"format": {"duration": "12"}, "streams": streams}))
                self.assertEqual(report["error_code"], "EXPORT_PROBE_FAILED")

    def test_invalid_or_duplicate_ffprobe_stream_indices_fail_closed(self):
        for indices in ((0, 0), (-1, 1), (True, 1), ("0", 1)):
            with self.subTest(indices=indices):
                streams = [{"codec_type": "video", "index": indices[0]},
                           {"codec_type": "audio", "index": indices[1]}]
                report = self.inspect_with_probe(json.dumps(
                    {"format": {"duration": "12"}, "streams": streams}))
                self.assertEqual(report["error_code"], "EXPORT_PROBE_FAILED")
        valid = self.inspect_with_probe(json.dumps({
            "format": {"duration": "12"},
            "streams": [{"codec_type": "video", "index": 0},
                        {"codec_type": "audio", "index": 3}]}))
        self.assertTrue(valid["passed"], valid)

    def test_nonfinite_filter_timestamps_do_not_enter_qa_report(self):
        huge = "9" * 400
        self.assertEqual(qa._signal_ranges(
            "black_start: 1 black_end: " + huge,
            "black_start", "black_end"), [])
        self.assertEqual(qa._signal_ranges(
            "silence_start: " + huge + " silence_end: 3",
            "silence_start", "silence_end"), [])
        self.assertEqual(qa._signal_ranges(
            "black_start: 1 black_end: 2", "black_start", "black_end"),
            [(1.0, 2.0)])

    def test_file_metadata_os_error_returns_structured_failure(self):
        with patch.object(qa.Path, "is_file", side_effect=OSError("access denied")):
            report = qa.inspect(str(self.video))
        self.assertEqual(report["error_code"], "EXPORT_VERIFY_FAILED")
        self.assertFalse(report["passed"])


    def test_rejects_impossible_or_incomplete_video_dimensions(self):
        for dimensions in (
            {"width": 0, "height": 1080},
            {"width": -1, "height": 1080},
            {"width": True, "height": 1080},
            {"width": "1920", "height": 1080},
            {"width": 1920},
            {"height": 1080},
            {"width": 16385, "height": 1080},
        ):
            with self.subTest(dimensions=dimensions):
                streams = [{"codec_type": "video", **dimensions},
                           {"codec_type": "audio"}]
                report = self.inspect_with_probe(json.dumps({
                    "format": {"duration": "12"}, "streams": streams}))
                self.assertEqual(report["error_code"], "EXPORT_PROBE_FAILED")
                self.assertFalse(report["passed"])
        valid = self.inspect_with_probe(json.dumps({
            "format": {"duration": "12"},
            "streams": [{"codec_type": "video", "width": 1920, "height": 1080},
                        {"codec_type": "audio"}]}))
        self.assertTrue(valid["passed"], valid)

    def test_rejects_unusable_audio_sample_rate_and_channels(self):
        for fields in (
            {"sample_rate": "0", "channels": 2},
            {"sample_rate": "-48000", "channels": 2},
            {"sample_rate": "48k", "channels": 2},
            {"sample_rate": 48000, "channels": 2},
            {"sample_rate": None, "channels": 2},
            {"sample_rate": "48000", "channels": 0},
            {"sample_rate": "48000", "channels": False},
            {"sample_rate": "48000", "channels": "2"},
        ):
            with self.subTest(fields=fields):
                report = self.inspect_with_probe(json.dumps({
                    "format": {"duration": "12"},
                    "streams": [{"codec_type": "video"},
                                {"codec_type": "audio", **fields}]}))
                self.assertEqual(report["error_code"], "EXPORT_PROBE_FAILED")
                self.assertFalse(report["passed"])
        valid = self.inspect_with_probe(json.dumps({
            "format": {"duration": "12"},
            "streams": [{"codec_type": "video"},
                        {"codec_type": "audio", "sample_rate": "48000", "channels": 2}]}))
        self.assertTrue(valid["passed"], valid)


if __name__ == "__main__":
    unittest.main()
