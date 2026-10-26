"""Regression tests for decode QA tail coverage and invalid subprocess output."""
import json
import subprocess
import sys
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
        # Test the real bounded capture path, not only its subprocess kwargs.
        command = [sys.executable, "-c",
                   "import sys;sys.stdout.buffer.write(b'\\xff');"
                   "sys.stderr.buffer.write(b'\\xfe')"]
        result = qa.run(command, 5)
        self.assertEqual(result.returncode, 0)
        self.assertEqual(result.stdout, "\ufffd")
        self.assertEqual(result.stderr, "\ufffd")

    def test_subprocess_output_is_bounded_even_when_child_writes_megabytes(self):
        command = [sys.executable, "-c",
                   "import sys;sys.stdout.buffer.write(b'A'*1000100);"
                   "sys.stderr.buffer.write(b'B'*1000010+b'END')"]
        result = qa.run(command, 5)
        self.assertEqual(result.returncode, 0)
        self.assertEqual(len(result.stdout), 1_000_001)
        self.assertEqual(result.stdout, "A" * 1_000_001)
        self.assertEqual(len(result.stderr), 1_000_000)
        self.assertTrue(result.stderr.endswith("END"))

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


    def test_rejects_boolean_and_ambiguous_duration_values(self):
        for duration in (True, False, [], {}, " 12", "12 ", "", None):
            with self.subTest(duration=duration):
                report = self.inspect_with_probe(json.dumps({
                    "format": {"duration": duration},
                    "streams": [{"codec_type": "video"}]}))
                self.assertEqual(report["error_code"], "EXPORT_PROBE_FAILED")
        valid = self.inspect_with_probe(json.dumps({
            "format": {"duration": "12.5"},
            "streams": [{"codec_type": "video"}]}))
        self.assertEqual(valid["duration_seconds"], 12.5)

    def test_rejects_nonstandard_json_constants_even_in_unused_metadata(self):
        for token in ("NaN", "Infinity", "-Infinity"):
            with self.subTest(token=token):
                report = self.inspect_with_probe(
                    '{"format":{"duration":"12","extra":' + token
                    + '},"streams":[{"codec_type":"video"}]}')
                self.assertEqual(report["error_code"], "EXPORT_PROBE_FAILED")

    def test_samples_every_video_and_audio_track_and_detects_secondary_failure(self):
        calls = []
        metadata = json.dumps({
            "format": {"duration": "12"},
            "streams": [
                {"codec_type": "video", "index": 0},
                {"codec_type": "video", "index": 2},
                {"codec_type": "audio", "index": 4},
                {"codec_type": "audio", "index": 7}]})

        def fake(command, timeout):
            if command[0] == "ffprobe":
                return subprocess.CompletedProcess(command, 0, metadata, "")
            mapping = command[command.index("-map") + 1]
            calls.append(mapping)
            return subprocess.CompletedProcess(
                command, 1 if mapping == "0:a:1" else 0, "",
                "broken secondary audio" if mapping == "0:a:1" else "")

        with patch.object(qa, "run", side_effect=fake):
            report = qa.inspect(str(self.video))
        self.assertFalse(report["passed"], report)
        self.assertEqual(set(calls), {"0:v:0", "0:v:1", "0:a:0", "0:a:1"})
        self.assertTrue(any(issue["code"] == "EXPORT_DECODE_FAILED"
                            and issue["stream_type"] == "audio"
                            and issue["stream_ordinal"] == 1
                            for issue in report["issues"]))

    def test_rejects_unbounded_number_of_decode_tracks(self):
        metadata = json.dumps({
            "format": {"duration": "12"},
            "streams": [{"codec_type": "video"}] +
                       [{"codec_type": "audio"} for _ in range(16)]})
        report = self.inspect_with_probe(metadata)
        self.assertEqual(report["error_code"], "EXPORT_STREAM_LIMIT")
        self.assertFalse(report["passed"])

    def test_rejects_media_replaced_during_decode(self):
        metadata = json.dumps({
            "format": {"duration": "12"},
            "streams": [{"codec_type": "video"}]})
        changed = False

        def fake(command, timeout):
            nonlocal changed
            if command[0] == "ffprobe":
                return subprocess.CompletedProcess(command, 0, metadata, "")
            if not changed:
                self.video.write_bytes(b"replacement video bytes" * 5)
                changed = True
            return subprocess.CompletedProcess(command, 0, "", "")

        with patch.object(qa, "run", side_effect=fake):
            report = qa.inspect(str(self.video))
        self.assertEqual(report["error_code"], "EXPORT_FILE_CHANGED")
        self.assertFalse(report["passed"])


    def test_negative_filter_start_timestamps_are_clamped(self):
        self.assertEqual(qa._signal_ranges(
            "silence_start: -1.5 silence_end: 2.0",
            "silence_start", "silence_end"), [(0.0, 2.0)])
        self.assertEqual(qa._signal_ranges(
            "black_start: -3 black_end: -1",
            "black_start", "black_end"), [])
        self.assertEqual(qa._signal_ranges(
            "black_start: 1.5 black_end: 2.5",
            "black_start", "black_end"), [(1.5, 2.5)])

    def test_deeply_nested_ffprobe_metadata_fails_closed(self):
        metadata = ('{"format":{"duration":"12"},'
                    '"streams":[{"codec_type":"video"}],"extra":'
                    + "[" * 1500 + "0" + "]" * 1500 + "}")
        report = self.inspect_with_probe(metadata)
        self.assertEqual(report["error_code"], "EXPORT_PROBE_FAILED")
        self.assertFalse(report["passed"])

    def test_submillisecond_samples_keep_nonzero_ffmpeg_time(self):
        calls = []
        metadata = json.dumps({"format": {"duration": "0.0004"},
                               "streams": [{"codec_type": "video"}]})

        def fake(command, timeout):
            calls.append(command)
            return subprocess.CompletedProcess(
                command, 0, metadata if command[0] == "ffprobe" else "", "")

        with patch.object(qa, "run", side_effect=fake):
            report = qa.inspect(str(self.video), require_audio=False)
        self.assertTrue(report["passed"], report)
        self.assertEqual(report["samples"][0]["duration_seconds"], 0.0004)
        ffmpeg = next(command for command in calls if command[0] == "ffmpeg")
        self.assertEqual(ffmpeg[ffmpeg.index("-t") + 1], "0.000400")

        too_short = self.inspect_with_probe(json.dumps({
            "format": {"duration": "0.0000004"},
            "streams": [{"codec_type": "video"}]}))
        self.assertEqual(too_short["error_code"], "EXPORT_PROBE_FAILED")

    def test_unknown_stream_types_do_not_silently_pass(self):
        malformed = self.inspect_with_probe(json.dumps({
            "format": {"duration": "12"},
            "streams": [{"codec_type": "video"}, {"codec_type": "audio"},
                        {"codec_type": "vide0"}]}))
        self.assertEqual(malformed["error_code"], "EXPORT_PROBE_FAILED")
        valid = self.inspect_with_probe(json.dumps({
            "format": {"duration": "12"},
            "streams": [{"codec_type": "video"}, {"codec_type": "audio"},
                        {"codec_type": "subtitle"}]}))
        self.assertTrue(valid["passed"], valid)

    def test_out_of_range_stream_indices_fail_closed(self):
        invalid = self.inspect_with_probe(json.dumps({
            "format": {"duration": "12"},
            "streams": [{"codec_type": "video", "index": 2147483648},
                        {"codec_type": "audio", "index": 1}]}))
        self.assertEqual(invalid["error_code"], "EXPORT_PROBE_FAILED")
        valid = self.inspect_with_probe(json.dumps({
            "format": {"duration": "12"},
            "streams": [{"codec_type": "video", "index": 2147483647},
                        {"codec_type": "audio", "index": 1}]}))
        self.assertTrue(valid["passed"], valid)




    def test_limits_non_av_streams_before_any_decode(self):
        metadata = json.dumps({
            "format": {"duration": "12"},
            "streams": [{"codec_type": "video"}] +
                       [{"codec_type": "subtitle"} for _ in range(128)]})
        calls = []

        def fake(command, timeout):
            calls.append(command[0])
            return subprocess.CompletedProcess(command, 0, metadata, "")

        with patch.object(qa, "run", side_effect=fake):
            report = qa.inspect(str(self.video), require_audio=False)
        self.assertEqual(report["error_code"], "EXPORT_STREAM_LIMIT")
        self.assertEqual(calls, ["ffprobe"])

    def test_rejects_out_of_range_audio_formats(self):
        for fields in ({"sample_rate": "768001", "channels": 2},
                       {"sample_rate": "48000", "channels": 65}):
            with self.subTest(fields=fields):
                report = self.inspect_with_probe(json.dumps({
                    "format": {"duration": "12"},
                    "streams": [{"codec_type": "video"},
                                {"codec_type": "audio", **fields}]}))
                self.assertEqual(report["error_code"], "EXPORT_PROBE_FAILED")
        valid = self.inspect_with_probe(json.dumps({
            "format": {"duration": "12"},
            "streams": [{"codec_type": "video"},
                        {"codec_type": "audio",
                         "sample_rate": "768000", "channels": 64}]}))
        self.assertTrue(valid["passed"], valid)

    def test_sampling_has_shared_deadline_across_tracks(self):
        metadata = json.dumps({
            "format": {"duration": "4"},
            "streams": [{"codec_type": "video"}, {"codec_type": "video"}]})
        ffmpeg_calls = []

        def fake(command, timeout):
            if command[0] == "ffprobe":
                return subprocess.CompletedProcess(command, 0, metadata, "")
            ffmpeg_calls.append(timeout)
            return subprocess.CompletedProcess(command, 0, "", "")

        with patch.object(qa, "run", side_effect=fake), \
                patch.object(qa, "time") as clock:
            # Create deadline; decode first track; second track is too late.
            clock.monotonic.side_effect = [0, 0, 121]
            report = qa.inspect(str(self.video), require_audio=False)
        self.assertEqual(report["error_code"], "EXPORT_DECODE_TIMEOUT")
        self.assertEqual(len(ffmpeg_calls), 1)
        self.assertLessEqual(ffmpeg_calls[0], 30)

    def test_full_decode_shares_budget_between_video_and_audio(self):
        metadata = json.dumps({
            "format": {"duration": "4"},
            "streams": [{"codec_type": "video"}, {"codec_type": "audio"}]})
        mappings = []

        def fake(command, timeout):
            if command[0] == "ffprobe":
                return subprocess.CompletedProcess(command, 0, metadata, "")
            mappings.append(command[command.index("-map") + 1])
            return subprocess.CompletedProcess(command, 0, "", "")

        with patch.object(qa, "run", side_effect=fake), \
                patch.object(qa, "time") as clock:
            # Sampling: deadline + 2 tracks; full: deadline + 2 tracks.
            clock.monotonic.side_effect = [0, 0, 0, 0, 0, 901]
            report = qa.inspect(str(self.video), full_decode=True)
        self.assertFalse(report["passed"], report)
        self.assertFalse(report["fully_decoded"])
        self.assertTrue(any(issue["code"] == "EXPORT_FULL_DECODE_TIMEOUT"
                            and issue["stream_type"] == "audio"
                            for issue in report["issues"]))
        self.assertEqual(mappings, ["0:v:0", "0:a:0", "0:v"])


    def test_rejects_ambiguous_ffprobe_duration_spellings(self):
        for duration in ("1_2", "+12", "１２", "12\\n", "0x1p0", "-1", "1e"):
            with self.subTest(duration=duration):
                report = self.inspect_with_probe(json.dumps({
                    "format": {"duration": duration},
                    "streams": [{"codec_type": "video"}]}))
                self.assertEqual(report["error_code"], "EXPORT_PROBE_FAILED")
        valid = self.inspect_with_probe(json.dumps({
            "format": {"duration": "1.25e1"},
            "streams": [{"codec_type": "video"}, {"codec_type": "audio"}]}))
        self.assertTrue(valid["passed"], valid)
        self.assertEqual(valid["duration_seconds"], 12.5)

    def test_rejects_uncheckably_long_video_without_decoding(self):
        for duration in ("604801", "1e308"):
            with self.subTest(duration=duration):
                calls = []
                metadata = json.dumps({
                    "format": {"duration": duration},
                    "streams": [{"codec_type": "video"}]})

                def fake(command, timeout):
                    calls.append(command[0])
                    return subprocess.CompletedProcess(command, 0, metadata, "")

                with patch.object(qa, "run", side_effect=fake):
                    report = qa.inspect(str(self.video), require_audio=False)
                self.assertEqual(report["error_code"], "EXPORT_DURATION_LIMIT")
                self.assertEqual(calls, ["ffprobe"])

if __name__ == "__main__":
    unittest.main()
