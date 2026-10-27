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

    def test_long_video_checks_first_middle_and_last_windows(self):
        self.check_windows(30.5, 8, [0.0, 11.25, 22.5])

    def test_exact_three_windows_checks_the_entire_duration(self):
        self.check_windows(24.0, 8, [0.0, 8.0, 16.0])

    def test_just_under_three_windows_does_not_duplicate_overlap(self):
        self.check_windows(23.5, 8, [0.0, 15.5])

    def test_subprocess_stdout_requires_utf8_but_stderr_is_diagnostic(self):
        # Metadata must never be silently altered by replacement decoding.
        invalid_stdout = [sys.executable, "-c",
                          "import sys;sys.stdout.buffer.write(b'\\xff')"]
        with self.assertRaises(UnicodeDecodeError):
            qa.run(invalid_stdout, 5)
        invalid_stderr = [sys.executable, "-c",
                          "import sys;sys.stdout.buffer.write(b'OK');"
                          "sys.stderr.buffer.write(b'\\xfe')"]
        result = qa.run(invalid_stderr, 5)
        self.assertEqual(result.stdout, "OK")
        self.assertEqual(result.stderr, "\ufffd")

    def test_subprocess_output_over_cap_is_rejected_not_silently_truncated(self):
        for stream in ("stdout", "stderr"):
            with self.subTest(stream=stream):
                command = [sys.executable, "-c",
                           "import sys;sys." + stream +
                           ".buffer.write(b'A'*1100000);sys." + stream + ".flush()"]
                with self.assertRaisesRegex(OSError, "exceeded the capture limit"):
                    qa.run(command, 5)

    def test_subprocess_output_at_caps_remains_usable(self):
        command = [sys.executable, "-c",
                   "import sys;sys.stdout.buffer.write(b'A'*1000001);"
                   "sys.stderr.buffer.write(b'B'*999997+b'END')"]
        result = qa.run(command, 5)
        self.assertEqual(result.returncode, 0)
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

    def test_json_exponent_overflow_is_rejected_even_in_unused_fields(self):
        prefix = ('{"format":{"duration":"8"},"streams":'
                  '[{"codec_type":"video"},{"codec_type":"audio"}],'
                  '"unused":{"value":')
        for value in ("1e9999", "-1e9999", "1e309"):
            with self.subTest(value=value):
                report = self.inspect_with_probe(prefix + value + "}}")
                self.assertEqual(report["error_code"], "EXPORT_PROBE_FAILED")
        valid = self.inspect_with_probe(prefix + "1.25e2}}")
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

    def test_rejects_pathological_frame_area_before_decoding(self):
        # Both dimensions individually pass the old 16K limit, but their
        # product would require enormous decoded frame allocations.
        metadata = json.dumps({
            "format": {"duration": "8"},
            "streams": [{"codec_type": "video",
                         "width": 16384, "height": 16384}]})
        self.assertEqual(self.inspect_with_probe(metadata)["error_code"],
                         "EXPORT_PROBE_FAILED")
        # Ordinary 8K UHD remains supported.
        valid = json.dumps({
            "format": {"duration": "8"},
            "streams": [{"codec_type": "video",
                         "width": 7680, "height": 4320},
                        {"codec_type": "audio"}]})
        self.assertTrue(self.inspect_with_probe(valid)["passed"])

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

    def test_invalid_utf8_ffprobe_is_structured_failure(self):
        with patch.object(qa, "run",
                          side_effect=UnicodeDecodeError(
                              "utf-8", b"\\xff", 0, 1, "invalid start byte")):
            report = qa.inspect(str(self.video))
        self.assertEqual(report["error_code"], "EXPORT_PROBE_FAILED")
        self.assertFalse(report["passed"])

    def test_probe_limit_counts_utf8_bytes_not_unicode_characters(self):
        # 400k characters occupy 1.2MB of UTF-8, exceeding the 1MB limit.
        metadata = json.dumps({
            "format": {"duration": "12"},
            "streams": [{"codec_type": "video"}],
            "extra": "あ" * 400_000}, ensure_ascii=False)
        self.assertLess(len(metadata), 1_000_000)
        self.assertGreater(len(metadata.encode("utf-8")), 1_000_000)
        report = self.inspect_with_probe(metadata)
        self.assertEqual(report["error_code"], "EXPORT_PROBE_FAILED")

    def test_filter_event_exponents_are_not_truncated(self):
        self.assertEqual(qa._signal_ranges(
            "black_start: 1e-2 black_end: 2e-2",
            "black_start", "black_end"), [(0.01, 0.02)])
        self.assertEqual(qa._signal_ranges(
            "black_start: 1e- black_end: 2",
            "black_start", "black_end"), [])

    def test_audio_peak_exponents_are_not_truncated(self):
        self.assertEqual(qa._audio_peak_db(
            "Peak level dB: -1e-2"), -0.01)
        self.assertIsNone(qa._audio_peak_db(
            "Peak level dB: 1e-"))


    def test_streaming_capture_rejects_large_output_without_deadlock(self):
        # Both pipes must be drained concurrently, and excessive diagnostics
        # must terminate the child rather than silently certify truncated QA.
        command = [sys.executable, "-c",
                   "import sys;"
                   "sys.stdout.buffer.write(b'A'*12000000);"
                   "sys.stderr.buffer.write(b'B'*12000000+b'END')"]
        with self.assertRaisesRegex(OSError, "exceeded the capture limit"):
            qa.run(command, 15)

    def test_streaming_capture_kills_timed_out_child(self):
        command = [sys.executable, "-c", "import time;time.sleep(5)"]
        with self.assertRaises(subprocess.TimeoutExpired):
            qa.run(command, 0.1)

    def test_open_black_and_silence_spans_reach_sample_eof(self):
        self.assertEqual(qa._signal_ranges(
            "black_start: 2.5", "black_start", "black_end", 8),
            [(2.5, 8)])
        self.assertEqual(qa._signal_ranges(
            "silence_start: -2", "silence_start", "silence_end", 8),
            [(0, 8)])
        # Without an observed window end, do not invent a closing timestamp.
        self.assertEqual(qa._signal_ranges(
            "black_start: 2.5", "black_start", "black_end"), [])

    def test_signal_ranges_never_extend_outside_decoded_window(self):
        self.assertEqual(qa._signal_ranges(
            "black_start: 7 black_end: 100",
            "black_start", "black_end", 8), [(7, 8)])
        self.assertEqual(qa._signal_ranges(
            "black_start: 9 black_end: 12",
            "black_start", "black_end", 8), [])
        self.assertEqual(qa._signal_ranges(
            "silence_start: -1 silence_end: 0.5",
            "silence_start", "silence_end", 8), [(0, 0.5)])

    def test_shorter_audio_track_warns_even_when_tail_decode_succeeds(self):
        metadata = json.dumps({
            "format": {"duration": "20"},
            "streams": [{"codec_type": "video", "duration": "20.0"},
                        {"codec_type": "audio", "start_time": "0.5",
                         "duration": "12.0"}]})
        result = self.inspect_with_probe(metadata)
        self.assertTrue(result["passed"], result)
        warnings = [i for i in result["issues"]
                    if i["code"] == "EXPORT_STREAM_ENDS_EARLY"]
        self.assertEqual(len(warnings), 1)
        self.assertEqual(warnings[0]["stream_type"], "audio")
        self.assertEqual(warnings[0]["missing_tail_seconds"], 7.5)

        within_tolerance = json.dumps({
            "format": {"duration": "20"},
            "streams": [{"codec_type": "video", "duration": "19.5"},
                        {"codec_type": "audio", "duration": "N/A"}]})
        report = self.inspect_with_probe(within_tolerance)
        self.assertFalse(any(i["code"] == "EXPORT_STREAM_ENDS_EARLY"
                             for i in report["issues"]))

    def test_invalid_optional_stream_times_are_rejected(self):
        for fields in ({"duration": "1_2"}, {"duration": "-1"},
                       {"duration": "1e999"}, {"start_time": "abc"},
                       {"start_time": True}):
            with self.subTest(fields=fields):
                metadata = json.dumps({
                    "format": {"duration": "20"},
                    "streams": [{"codec_type": "video", **fields}]})
                result = self.inspect_with_probe(metadata)
                self.assertEqual(result["error_code"], "EXPORT_PROBE_FAILED")


    def test_reports_validated_video_resolution_and_fractional_fps(self):
        metadata = json.dumps({
            "format": {"duration": "12"},
            "streams": [
                {"codec_type": "video", "width": 1920, "height": 1080,
                 "avg_frame_rate": "30000/1001"},
                {"codec_type": "video", "width": 1280, "height": 720,
                 "avg_frame_rate": "25/1"},
                {"codec_type": "audio", "sample_rate": "48000", "channels": 2}]})
        report = self.inspect_with_probe(metadata)
        self.assertTrue(report["passed"], report)
        self.assertEqual(report["resolution"], "1920x1080")
        self.assertEqual(report["fps"], 29.97)
        self.assertEqual(report["video_details"], [
            {"stream_ordinal": 0, "width": 1920, "height": 1080, "fps": 29.97},
            {"stream_ordinal": 1, "width": 1280, "height": 720, "fps": 25.0}])
        self.assertEqual(report["audio_details"], [
            {"stream_ordinal": 0, "sample_rate": 48000, "channels": 2}])

    def test_rejects_malformed_or_impossible_video_frame_rates(self):
        for rate in ("30", "29.97", "1/0", "1001/1", "1/1000000000",
                     "1_0/1", "３0/1", True, 30, "-1/1"):
            with self.subTest(rate=rate):
                metadata = json.dumps({
                    "format": {"duration": "12"},
                    "streams": [{"codec_type": "video", "avg_frame_rate": rate}]})
                self.assertEqual(self.inspect_with_probe(metadata)["error_code"],
                                 "EXPORT_PROBE_FAILED")

    def test_unknown_video_fps_does_not_claim_measured_frame_rate(self):
        for rate in ("0/0", "0/1", "N/A"):
            with self.subTest(rate=rate):
                metadata = json.dumps({
                    "format": {"duration": "12"},
                    "streams": [{"codec_type": "video", "avg_frame_rate": rate},
                                {"codec_type": "audio"}]})
                report = self.inspect_with_probe(metadata)
                self.assertTrue(report["passed"], report)
                self.assertIsNone(report["fps"])
                self.assertIsNone(report["video_details"][0]["fps"])

    def test_reports_missing_optional_media_properties_as_unknown(self):
        metadata = json.dumps({
            "format": {"duration": "12"},
            "streams": [{"codec_type": "video"}, {"codec_type": "audio"}]})
        report = self.inspect_with_probe(metadata)
        self.assertTrue(report["passed"], report)
        self.assertIsNone(report["resolution"])
        self.assertIsNone(report["fps"])
        self.assertEqual(report["audio_details"][0]["sample_rate"], None)
        self.assertEqual(report["audio_details"][0]["channels"], None)

    def test_frozen_video_frames_are_reported_without_failing_export(self):
        metadata = json.dumps({
            "format": {"duration": "8"},
            "streams": [{"codec_type": "video"}, {"codec_type": "audio"}]})
        video_commands = []

        def fake(command, timeout):
            if command[0] == "ffprobe":
                return subprocess.CompletedProcess(command, 0, metadata, "")
            if "0:v:0" in command:
                video_commands.append(command)
                return subprocess.CompletedProcess(
                    command, 0, "", "freeze_start: 2 freeze_end: 4")
            return subprocess.CompletedProcess(command, 0, "", "")

        with patch.object(qa, "run", side_effect=fake):
            report = qa.inspect(str(self.video))
        self.assertTrue(report["passed"], report)
        self.assertEqual(len(video_commands), 1)
        self.assertIn("freezedetect=n=-60dB:d=2",
                      video_commands[0][video_commands[0].index("-vf") + 1])
        self.assertEqual(report["frozen_frames"][0]["duration_seconds"], 2)
        self.assertTrue(any(i["code"] == "EXPORT_SAMPLE_FROZEN_FRAMES"
                            for i in report["issues"]))

    def test_freeze_without_end_event_extends_to_sample_end(self):
        metadata = json.dumps({
            "format": {"duration": "8"},
            "streams": [{"codec_type": "video"}]})

        def fake(command, timeout):
            return subprocess.CompletedProcess(
                command, 0, metadata if command[0] == "ffprobe" else "",
                "freeze_start: 3" if command[0] == "ffmpeg" else "")

        with patch.object(qa, "run", side_effect=fake):
            report = qa.inspect(str(self.video), require_audio=False)
        self.assertTrue(report["passed"], report)
        self.assertEqual(report["frozen_frames"][0]["filter_end_seconds"], 8)

    def test_warns_on_late_start_and_track_extending_past_container(self):
        metadata = json.dumps({
            "format": {"duration": "20"},
            "streams": [{"codec_type": "video", "start_time": "2",
                         "duration": "20"},
                        {"codec_type": "audio", "start_time": "0",
                         "duration": "20"}]})
        report = self.inspect_with_probe(metadata)
        self.assertTrue(report["passed"], report)
        codes = [issue["code"] for issue in report["issues"]]
        self.assertIn("EXPORT_STREAM_STARTS_LATE", codes)
        self.assertIn("EXPORT_STREAM_EXCEEDS_CONTAINER", codes)
        self.assertEqual(next(issue["missing_head_seconds"]
                              for issue in report["issues"]
                              if issue["code"] == "EXPORT_STREAM_STARTS_LATE"), 2.0)

    def test_rejects_stream_start_beyond_container_end(self):
        metadata = json.dumps({
            # Explicit zero origin proves this is a late track, not a
            # container with nonzero presentation timestamps.
            "format": {"duration": "20", "start_time": "0"},
            "streams": [{"codec_type": "video", "start_time": "23"}]})
        self.assertEqual(self.inspect_with_probe(metadata)["error_code"],
                         "EXPORT_PROBE_FAILED")


    def test_failed_full_decode_does_not_claim_full_verification_scope(self):
        metadata = json.dumps({
            "format": {"duration": "4"},
            "streams": [{"codec_type": "video"}, {"codec_type": "audio"}]})

        def fake(command, timeout):
            if command[0] == "ffprobe":
                return subprocess.CompletedProcess(command, 0, metadata, "")
            is_full = "-ss" not in command
            return subprocess.CompletedProcess(
                command, 1 if is_full else 0, "", "decode failed" if is_full else "")

        with patch.object(qa, "run", side_effect=fake):
            report = qa.inspect(str(self.video), full_decode=True)
        self.assertFalse(report["passed"], report)
        self.assertFalse(report["fully_decoded"])
        self.assertEqual(report["verification_scope"], "incomplete")

    def test_successful_full_decode_and_sampling_have_distinct_scopes(self):
        metadata = json.dumps({
            "format": {"duration": "4"},
            "streams": [{"codec_type": "video"}, {"codec_type": "audio"}]})

        def fake(command, timeout):
            return subprocess.CompletedProcess(
                command, 0, metadata if command[0] == "ffprobe" else "", "")

        with patch.object(qa, "run", side_effect=fake):
            full = qa.inspect(str(self.video), full_decode=True)
            sampled = qa.inspect(str(self.video), full_decode=False)
        self.assertTrue(full["fully_decoded"])
        self.assertEqual(full["verification_scope"], "full_audio_video_decode")
        self.assertFalse(sampled["fully_decoded"])
        self.assertEqual(sampled["verification_scope"], "sampled_audio_video_decode")


    def test_multitrack_video_and_audio_format_mismatches_are_warnings(self):
        metadata = json.dumps({
            "format": {"duration": "8"},
            "streams": [
                {"codec_type": "video", "width": 1920, "height": 1080,
                 "avg_frame_rate": "30/1"},
                {"codec_type": "video", "width": 1280, "height": 720,
                 "avg_frame_rate": "24/1"},
                {"codec_type": "audio", "sample_rate": "48000", "channels": 2},
                {"codec_type": "audio", "sample_rate": "44100", "channels": 1}]})
        report = self.inspect_with_probe(metadata)
        self.assertTrue(report["passed"], report)
        warnings = {i["code"]: i for i in report["issues"]}
        self.assertEqual(warnings["EXPORT_VIDEO_TRACK_RESOLUTION_MISMATCH"]
                         ["resolutions"], ["1280x720", "1920x1080"])
        self.assertEqual(warnings["EXPORT_VIDEO_TRACK_FPS_MISMATCH"]
                         ["frame_rates"], [24.0, 30.0])
        self.assertEqual(warnings["EXPORT_AUDIO_TRACK_SAMPLE_RATE_MISMATCH"]
                         ["sample_rates"], [44100, 48000])
        self.assertEqual(warnings["EXPORT_AUDIO_TRACK_CHANNEL_MISMATCH"]
                         ["channel_counts"], [1, 2])

    def test_matching_and_unknown_track_formats_do_not_create_mismatch_warnings(self):
        metadata = json.dumps({
            "format": {"duration": "8"},
            "streams": [
                {"codec_type": "video", "width": 1920, "height": 1080,
                 "avg_frame_rate": "30/1"},
                {"codec_type": "video", "width": 1920, "height": 1080,
                 "avg_frame_rate": "N/A"},
                {"codec_type": "audio", "sample_rate": "48000", "channels": 2},
                {"codec_type": "audio", "channels": 2}]})
        report = self.inspect_with_probe(metadata)
        self.assertTrue(report["passed"], report)
        self.assertFalse(any("MISMATCH" in i["code"] for i in report["issues"]))


    def test_nonzero_container_origin_does_not_invent_track_gaps(self):
        metadata = json.dumps({
            "format": {"duration": "20", "start_time": "10.5"},
            "streams": [{"codec_type": "video", "start_time": "10.5",
                         "duration": "20"},
                        {"codec_type": "audio", "start_time": "10.5",
                         "duration": "20"}]})
        report = self.inspect_with_probe(metadata)
        self.assertTrue(report["passed"], report)
        self.assertFalse(any(i["code"] in (
            "EXPORT_STREAM_STARTS_LATE", "EXPORT_STREAM_EXCEEDS_CONTAINER",
            "EXPORT_STREAM_ENDS_EARLY") for i in report["issues"]))

    def test_track_delay_is_relative_to_container_origin(self):
        metadata = json.dumps({
            "format": {"duration": "20", "start_time": "10"},
            "streams": [{"codec_type": "video", "start_time": "12",
                         "duration": "18"},
                        {"codec_type": "audio", "start_time": "10",
                         "duration": "20"}]})
        report = self.inspect_with_probe(metadata)
        self.assertTrue(report["passed"], report)
        warning = next(i for i in report["issues"]
                       if i["code"] == "EXPORT_STREAM_STARTS_LATE")
        self.assertEqual(warning["missing_head_seconds"], 2.0)
        self.assertFalse(any(i["code"] == "EXPORT_STREAM_EXCEEDS_CONTAINER"
                             for i in report["issues"]))

    def test_invalid_container_start_time_fails_closed(self):
        for start in (True, 5, "NaN", "1_0", "1e9999"):
            with self.subTest(start=start):
                metadata = json.dumps({
                    "format": {"duration": "20", "start_time": start},
                    "streams": [{"codec_type": "video"}]})
                report = self.inspect_with_probe(metadata)
                self.assertEqual(report["error_code"], "EXPORT_PROBE_FAILED")

    def test_signal_warnings_include_absolute_video_timestamps(self):
        metadata = json.dumps({
            "format": {"duration": "12.5"},
            "streams": [{"codec_type": "video"}, {"codec_type": "audio"}]})
        def fake(command, timeout):
            if command[0] == "ffprobe":
                return subprocess.CompletedProcess(command, 0, metadata, "")
            log = ("black_start: 1 black_end: 3 freeze_start: 4 freeze_end: 6"
                   if "0:v:0" in command else "silence_start: 2 silence_end: 4")
            return subprocess.CompletedProcess(command, 0, "", log)
        with patch.object(qa, "run", side_effect=fake):
            report = qa.inspect(str(self.video))
        self.assertTrue(report["passed"], report)
        for key, relative_start, relative_end in (
                ("black_frames", 1, 3), ("frozen_frames", 4, 6),
                ("long_silences", 2, 4)):
            event = next(e for e in report[key] if e["sample_start_seconds"] == 4.5)
            self.assertEqual(event["video_start_seconds"], 4.5 + relative_start)
            self.assertEqual(event["video_end_seconds"], 4.5 + relative_end)

    def test_file_identity_detects_restored_mtime_with_changed_ctime(self):
        from types import SimpleNamespace
        original = SimpleNamespace(st_dev=1, st_ino=2, st_size=100,
                                   st_mtime_ns=500, st_ctime_ns=700)
        changed = SimpleNamespace(st_dev=1, st_ino=2, st_size=100,
                                  st_mtime_ns=500, st_ctime_ns=701)
        self.assertNotEqual(qa._file_identity(original), qa._file_identity(changed))


    def test_missing_format_origin_uses_earliest_stream_timestamp(self):
        # Some containers omit format.start_time even when all tracks have
        # nonzero presentation timestamps. Their first track is not delayed.
        metadata = json.dumps({
            "format": {"duration": "20"},
            "streams": [{"codec_type": "video", "start_time": "10",
                         "duration": "20"},
                        {"codec_type": "audio", "start_time": "12",
                         "duration": "18"}]})
        report = self.inspect_with_probe(metadata)
        self.assertTrue(report["passed"], report)
        late = [issue for issue in report["issues"]
                if issue["code"] == "EXPORT_STREAM_STARTS_LATE"]
        self.assertEqual(len(late), 1)
        self.assertEqual(late[0]["stream_type"], "audio")
        self.assertEqual(late[0]["missing_head_seconds"], 2.0)
        self.assertFalse(any(issue["code"] in (
            "EXPORT_STREAM_ENDS_EARLY", "EXPORT_STREAM_EXCEEDS_CONTAINER")
            for issue in report["issues"]))

    def test_incomplete_track_timestamps_do_not_shift_container_origin(self):
        metadata = json.dumps({
            "format": {"duration": "20"},
            "streams": [{"codec_type": "video", "duration": "20"},
                        {"codec_type": "audio", "start_time": "0.5",
                         "duration": "12"}]})
        report = self.inspect_with_probe(metadata)
        self.assertTrue(report["passed"], report)
        warning = next(issue for issue in report["issues"]
                       if issue["code"] == "EXPORT_STREAM_ENDS_EARLY")
        self.assertEqual(warning["missing_tail_seconds"], 7.5)

    def test_missing_format_origin_without_track_timestamps_defaults_to_zero(self):
        metadata = json.dumps({
            "format": {"duration": "8"},
            "streams": [{"codec_type": "video"}, {"codec_type": "audio"}]})
        report = self.inspect_with_probe(metadata)
        self.assertTrue(report["passed"], report)
        self.assertFalse(any("STARTS" in issue["code"] for issue in report["issues"]))

    def test_stream_origin_subtraction_overflow_fails_closed(self):
        for origin, stream_start in (("1e308", "-1e308"),
                                     ("-1e308", "1e308")):
            with self.subTest(origin=origin):
                metadata = json.dumps({
                    "format": {"duration": "20", "start_time": origin},
                    "streams": [{"codec_type": "video",
                                 "start_time": stream_start, "duration": "20"}]})
                report = self.inspect_with_probe(metadata)
                self.assertEqual(report["error_code"], "EXPORT_PROBE_FAILED")

    def test_large_negative_stream_preroll_is_reported_not_rejected(self):
        metadata = json.dumps({
            "format": {"duration": "20", "start_time": "10"},
            "streams": [{"codec_type": "video", "start_time": "8",
                         "duration": "22"},
                        {"codec_type": "audio", "start_time": "10",
                         "duration": "20"}]})
        report = self.inspect_with_probe(metadata)
        self.assertTrue(report["passed"], report)
        early = [issue for issue in report["issues"]
                 if issue["code"] == "EXPORT_STREAM_STARTS_BEFORE_CONTAINER"]
        self.assertEqual(len(early), 1)
        self.assertEqual(early[0]["stream_type"], "video")
        self.assertEqual(early[0]["preroll_seconds"], 2.0)
        self.assertFalse(any(issue["code"] == "EXPORT_STREAM_EXCEEDS_CONTAINER"
                             for issue in report["issues"]))

    def test_unknown_track_origin_does_not_invent_end_gap(self):
        # A track with an unknown start might occupy the final 12 seconds.
        # Duration alone cannot establish that it ends before the container.
        metadata = json.dumps({
            "format": {"duration": "20", "start_time": "10"},
            "streams": [{"codec_type": "video", "start_time": "10",
                         "duration": "20"},
                        {"codec_type": "audio", "duration": "12"}]})
        report = self.inspect_with_probe(metadata)
        self.assertTrue(report["passed"], report)
        self.assertFalse(any(i["code"] in (
            "EXPORT_STREAM_ENDS_EARLY", "EXPORT_STREAM_EXCEEDS_CONTAINER",
            "EXPORT_STREAM_STARTS_LATE") for i in report["issues"]))

    def test_overlapping_sample_windows_do_not_repeat_same_signal(self):
        metadata = json.dumps({
            "format": {"duration": "12"},
            "streams": [{"codec_type": "video"}, {"codec_type": "audio"}]})

        def fake(command, timeout):
            if command[0] == "ffprobe":
                return subprocess.CompletedProcess(command, 0, metadata, "")
            start = float(command[command.index("-ss") + 1])
            relative_start, relative_end = 5 - start, 7 - start
            if command[command.index("-map") + 1] == "0:v:0":
                log = (f"black_start: {relative_start} black_end: {relative_end}\n"
                       f"freeze_start: {relative_start} freeze_end: {relative_end}\n")
            else:
                log = f"silence_start: {relative_start} silence_end: {relative_end}\n"
            return subprocess.CompletedProcess(command, 0, "", log)

        with patch.object(qa, "run", side_effect=fake):
            report = qa.inspect(str(self.video))
        self.assertTrue(report["passed"], report)
        self.assertEqual(len(report["samples"]), 2)
        for key in ("black_frames", "frozen_frames", "long_silences"):
            self.assertEqual(len(report[key]), 1, (key, report[key]))
            self.assertEqual(report[key][0]["video_start_seconds"], 5.0)
            self.assertEqual(report[key][0]["video_end_seconds"], 7.0)
        self.assertEqual(sum(i["code"].startswith("EXPORT_SAMPLE_") for i
                             in report["issues"]), 3)

    def test_partially_clipped_sample_signals_merge_and_update_issues(self):
        # One physical defect spans both 8-second windows, but each ffmpeg
        # invocation can report only the portion inside its own window.
        metadata = json.dumps({
            "format": {"duration": "12"},
            "streams": [{"codec_type": "video"}, {"codec_type": "audio"}]})

        def fake(command, timeout):
            if command[0] == "ffprobe":
                return subprocess.CompletedProcess(command, 0, metadata, "")
            start = float(command[command.index("-ss") + 1])
            relative_start, relative_end = 5 - start, 11 - start
            if "0:v:0" in command:
                log = (f"black_start: {relative_start} black_end: {relative_end}\n"
                       f"freeze_start: {relative_start} freeze_end: {relative_end}\n")
            else:
                log = f"silence_start: {relative_start} silence_end: {relative_end}\n"
            return subprocess.CompletedProcess(command, 0, "", log)

        with patch.object(qa, "run", side_effect=fake):
            report = qa.inspect(str(self.video))
        self.assertTrue(report["passed"], report)
        for key, code in (("black_frames", "EXPORT_SAMPLE_BLACK_FRAMES"),
                          ("frozen_frames", "EXPORT_SAMPLE_FROZEN_FRAMES"),
                          ("long_silences", "EXPORT_SAMPLE_LONG_SILENCE")):
            events = report[key]
            matching_issues = [issue for issue in report["issues"]
                               if issue["code"] == code]
            self.assertEqual(len(events), 1, (key, events))
            self.assertEqual(len(matching_issues), 1, matching_issues)
            for record in (events[0], matching_issues[0]):
                self.assertEqual(record["video_start_seconds"], 5)
                self.assertEqual(record["video_end_seconds"], 11)
                self.assertEqual(record["duration_seconds"], 6)

    def test_signal_merge_bridges_two_previous_intervals(self):
        # Two detections from the first window are connected by a later
        # detection. Coalescing must remove the redundant event AND QA issue.
        metadata = json.dumps({
            "format": {"duration": "12"},
            "streams": [{"codec_type": "video"}]})

        def fake(command, timeout):
            if command[0] == "ffprobe":
                return subprocess.CompletedProcess(command, 0, metadata, "")
            start = float(command[command.index("-ss") + 1])
            log = ("black_start: 5 black_end: 6\n"
                   "black_start: 7 black_end: 8\n" if start == 0
                   else "black_start: 1.5 black_end: 3.5\n")
            return subprocess.CompletedProcess(command, 0, "", log)

        with patch.object(qa, "run", side_effect=fake):
            report = qa.inspect(str(self.video), require_audio=False)
        self.assertTrue(report["passed"], report)
        self.assertEqual(len(report["black_frames"]), 1, report["black_frames"])
        event = report["black_frames"][0]
        self.assertEqual((event["video_start_seconds"], event["video_end_seconds"],
                          event["duration_seconds"]), (5, 8, 3))
        matching = [issue for issue in report["issues"]
                    if issue["code"] == "EXPORT_SAMPLE_BLACK_FRAMES"]
        self.assertEqual(len(matching), 1, matching)
        self.assertEqual(matching[0]["video_end_seconds"], 8)
        self.assertEqual(matching[0]["duration_seconds"], 3)

    def test_stream_end_overflow_is_invalid_metadata(self):
        # Although the container itself exceeds the duration policy, malformed
        # stream arithmetic must never be propagated as a valid finite value.
        metadata = json.dumps({
            "format": {"duration": "1e308", "start_time": "0"},
            "streams": [{"codec_type": "video", "start_time": "1e308",
                         "duration": "1e308"}]})
        report = self.inspect_with_probe(metadata)
        self.assertEqual(report["error_code"], "EXPORT_PROBE_FAILED")


    def test_distinct_frame_rate_ratios_do_not_collapse_after_rounding(self):
        metadata = json.dumps({
            "format": {"duration": "8"},
            "streams": [{"codec_type": "video", "avg_frame_rate": "30000/1001"},
                        {"codec_type": "video", "avg_frame_rate": "2997/100"},
                        {"codec_type": "audio"}]})
        report = self.inspect_with_probe(metadata)
        self.assertTrue(report["passed"], report)
        mismatches = [issue for issue in report["issues"]
                      if issue["code"] == "EXPORT_VIDEO_TRACK_FPS_MISMATCH"]
        self.assertEqual(len(mismatches), 1, report)
        self.assertEqual(mismatches[0]["frame_rates"], [29.97])
        self.assertEqual(mismatches[0]["frame_rate_ratios"],
                         ["2997/100", "30000/1001"])
        # Equivalent fractions must not trigger a spurious mismatch.
        equivalent = json.dumps({
            "format": {"duration": "8"},
            "streams": [{"codec_type": "video", "avg_frame_rate": "30000/1001"},
                        {"codec_type": "video", "avg_frame_rate": "60000/2002"},
                        {"codec_type": "audio"}]})
        report = self.inspect_with_probe(equivalent)
        self.assertTrue(report["passed"], report)
        self.assertFalse(any(issue["code"] == "EXPORT_VIDEO_TRACK_FPS_MISMATCH"
                             for issue in report["issues"]))

    def test_video_only_full_decode_does_not_claim_audio_was_decoded(self):
        metadata = json.dumps({
            "format": {"duration": "4"},
            "streams": [{"codec_type": "video", "avg_frame_rate": "30/1"}]})
        commands = []

        def fake(command, timeout):
            commands.append(command)
            return subprocess.CompletedProcess(
                command, 0, metadata if command[0] == "ffprobe" else "", "")

        with patch.object(qa, "run", side_effect=fake):
            report = qa.inspect(str(self.video), require_audio=False,
                                full_decode=True)
        self.assertTrue(report["passed"], report)
        self.assertTrue(report["fully_decoded"])
        self.assertEqual(report["verification_scope"], "full_video_decode")
        with patch.object(qa, "run", side_effect=fake):
            sampled = qa.inspect(str(self.video), require_audio=False)
        self.assertTrue(sampled["passed"], sampled)
        self.assertFalse(sampled["fully_decoded"])
        self.assertEqual(sampled["verification_scope"], "sampled_video_decode")
        self.assertFalse(any("0:a" in " ".join(command)
                             for command in commands if command[0] == "ffmpeg"))

    def test_skips_empty_windows_for_intentionally_delayed_tracks(self):
        metadata = json.dumps({
            "format": {"duration": "30", "start_time": "0"},
            "streams": [{"codec_type": "video", "start_time": "0", "duration": "30"},
                        {"codec_type": "audio", "start_time": "10", "duration": "15"}]})
        decoded_windows = []

        def fake(command, timeout):
            if command[0] == "ffprobe":
                return subprocess.CompletedProcess(command, 0, metadata, "")
            stream = command[command.index("-map") + 1]
            start = float(command[command.index("-ss") + 1])
            decoded_windows.append((stream, start))
            # An actual FFmpeg invocation for the empty head would fail.
            return subprocess.CompletedProcess(
                command, 1 if stream == "0:a:0" and start == 0 else 0, "", "")

        with patch.object(qa, "run", side_effect=fake):
            report = qa.inspect(str(self.video))
        self.assertTrue(report["passed"], report)
        self.assertNotIn(("0:a:0", 0.0), decoded_windows)
        self.assertIn(("0:a:0", 11.0), decoded_windows)
        self.assertIn(("0:a:0", 22.0), decoded_windows)
        self.assertIn("EXPORT_STREAM_STARTS_LATE",
                      [issue["code"] for issue in report["issues"]])

    def test_short_track_between_regular_windows_gets_targeted_sample(self):
        # Default windows for 30 seconds are 0-8, 11-19 and 22-30.
        # The one-second audio track at 9-10 seconds needs its own sample.
        metadata = json.dumps({
            "format": {"duration": "30", "start_time": "0"},
            "streams": [{"codec_type": "video", "start_time": "0",
                         "duration": "30"},
                        {"codec_type": "audio", "start_time": "9",
                         "duration": "1"}]})
        decoded = []

        def fake(command, timeout):
            if command[0] == "ffprobe":
                return subprocess.CompletedProcess(command, 0, metadata, "")
            if "-ss" in command:
                decoded.append((command[command.index("-map") + 1],
                                float(command[command.index("-ss") + 1])))
            return subprocess.CompletedProcess(command, 0, "", "")

        with patch.object(qa, "run", side_effect=fake):
            report = qa.inspect(str(self.video))
        self.assertTrue(report["passed"], report)
        self.assertEqual(len(report["samples"]), 4)
        audio_windows = [start for kind, start in decoded if kind == "0:a:0"]
        self.assertEqual(len(audio_windows), 1)
        self.assertLessEqual(audio_windows[0], 9)
        self.assertGreater(audio_windows[0] + 8, 10)
        self.assertFalse(any(issue["code"] == "EXPORT_STREAM_NOT_SAMPLED"
                             for issue in report["issues"]))

    def test_track_outside_container_remains_uncertified(self):
        # A track whose entire timestamp range is outside the container
        # cannot be made valid by inventing an extra sample window.
        metadata = json.dumps({
            "format": {"duration": "30", "start_time": "0"},
            "streams": [{"codec_type": "video", "start_time": "0",
                         "duration": "30"},
                        {"codec_type": "audio", "start_time": "30.2",
                         "duration": "0.2"}]})
        report = self.inspect_with_probe(metadata)
        self.assertFalse(report["passed"])
        self.assertIn("EXPORT_STREAM_NOT_SAMPLED",
                      [issue["code"] for issue in report["issues"]])

    def test_invalid_ffmpeg_sample_stdout_is_structured_failure(self):
        metadata = json.dumps({
            "format": {"duration": "4"},
            "streams": [{"codec_type": "video"}]})

        def fake(command, timeout):
            if command[0] == "ffprobe":
                return subprocess.CompletedProcess(command, 0, metadata, "")
            raise UnicodeDecodeError("utf-8", b"\\xff", 0, 1, "invalid start byte")

        with patch.object(qa, "run", side_effect=fake):
            report = qa.inspect(str(self.video), require_audio=False)
        self.assertFalse(report["passed"])
        self.assertEqual(report["error_code"], "EXPORT_DECODE_FAILED")

    def test_invalid_ffmpeg_full_decode_stdout_is_reported_not_raised(self):
        metadata = json.dumps({
            "format": {"duration": "4"},
            "streams": [{"codec_type": "video"}]})

        def fake(command, timeout):
            if command[0] == "ffprobe":
                return subprocess.CompletedProcess(command, 0, metadata, "")
            if "-ss" not in command:
                raise UnicodeDecodeError("utf-8", b"\\xff", 0, 1, "invalid start byte")
            return subprocess.CompletedProcess(command, 0, "", "")

        with patch.object(qa, "run", side_effect=fake):
            report = qa.inspect(str(self.video), require_audio=False, full_decode=True)
        self.assertFalse(report["passed"], report)
        self.assertFalse(report["fully_decoded"])
        self.assertEqual(report["verification_scope"], "incomplete")
        self.assertIn("EXPORT_FULL_DECODE_FAILED",
                      [issue["code"] for issue in report["issues"]])

    def test_partially_indexed_stream_table_is_rejected(self):
        # ffprobe normally indexes every stream. A mixed table may hide a
        # missing track or make the ordinal-to-stream mapping ambiguous.
        metadata = json.dumps({
            "format": {"duration": "12"},
            "streams": [{"codec_type": "video", "index": 0},
                        {"codec_type": "audio"}]})
        report = self.inspect_with_probe(metadata)
        self.assertFalse(report["passed"], report)
        self.assertEqual(report["error_code"], "EXPORT_PROBE_FAILED")
        # Legacy mocks that omit all indices remain supported.
        legacy = json.dumps({
            "format": {"duration": "12"},
            "streams": [{"codec_type": "video"}, {"codec_type": "audio"}]})
        self.assertTrue(self.inspect_with_probe(legacy)["passed"])

    def test_zero_duration_audio_or_video_stream_cannot_pass(self):
        # A reported empty stream must not pass merely because a mock or
        # permissive decoder returns exit code zero for a sample.
        for kind in ("video", "audio"):
            with self.subTest(kind=kind):
                metadata = json.dumps({
                    "format": {"duration": "12"},
                    "streams": [{"codec_type": "video", **(
                        {"duration": "0"} if kind == "video" else {})},
                        {"codec_type": "audio", **(
                            {"duration": "0.000"} if kind == "audio" else {})}]})
                report = self.inspect_with_probe(metadata)
                self.assertFalse(report["passed"], report)
                self.assertEqual(report["error_code"], "EXPORT_PROBE_FAILED")
        # Unknown durations must not be mistaken for known-empty tracks.
        unknown = json.dumps({
            "format": {"duration": "12"},
            "streams": [{"codec_type": "video", "duration": "N/A"},
                        {"codec_type": "audio"}]})
        self.assertTrue(self.inspect_with_probe(unknown)["passed"])

    def test_open_ended_signals_stop_at_known_track_end(self):
        # The file lasts 12s but both tracks end at 6s. A filter that emits
        # *_start at 5s without *_end must not claim the defect lasts to 8s
        # or 12s merely because those are the requested sample boundaries.
        metadata = json.dumps({
            "format": {"duration": "12", "start_time": "0"},
            "streams": [{"codec_type": "video", "start_time": "0",
                         "duration": "6"},
                        {"codec_type": "audio", "start_time": "0",
                         "duration": "6"}]})

        def fake(command, timeout):
            if command[0] == "ffprobe":
                return subprocess.CompletedProcess(command, 0, metadata, "")
            start = float(command[command.index("-ss") + 1])
            relative = 5 - start
            if "0:v:0" in command:
                log = (f"black_start: {relative}\n"
                       f"freeze_start: {relative}\n")
            else:
                log = f"silence_start: {relative}\n"
            return subprocess.CompletedProcess(command, 0, "", log)

        with patch.object(qa, "run", side_effect=fake):
            report = qa.inspect(str(self.video))
        self.assertTrue(report["passed"], report)
        for key in ("black_frames", "frozen_frames", "long_silences"):
            self.assertEqual(len(report[key]), 1, report[key])
            self.assertEqual(report[key][0]["video_start_seconds"], 5.0)
            self.assertEqual(report[key][0]["video_end_seconds"], 6.0)
            self.assertEqual(report[key][0]["duration_seconds"], 1.0)

    def test_declared_ffprobe_stream_count_must_match_actual_inventory(self):
        # A truncated stream table could otherwise hide a broken secondary
        # track while the QA report claims to have verified every stream.
        for count in (0, 1, 3, -1, True, "2"):
            with self.subTest(count=count):
                metadata = json.dumps({
                    "format": {"duration": "12", "nb_streams": count},
                    "streams": [{"codec_type": "video"}, {"codec_type": "audio"}]})
                self.assertEqual(self.inspect_with_probe(metadata)["error_code"],
                                 "EXPORT_PROBE_FAILED")
        valid = json.dumps({
            "format": {"duration": "12", "nb_streams": 2},
            "streams": [{"codec_type": "video"}, {"codec_type": "audio"}]})
        self.assertTrue(self.inspect_with_probe(valid)["passed"])

    def test_reordered_stream_indices_cannot_misattribute_track_diagnostics(self):
        # FFmpeg maps 0:v:N / 0:a:N by ascending stream index, not arbitrary
        # JSON array order. Sparse ascending indices remain valid.
        for indices in ((3, 1), (2, 0)):
            with self.subTest(indices=indices):
                metadata = json.dumps({
                    "format": {"duration": "12"},
                    "streams": [{"codec_type": "video", "index": indices[0]},
                                {"codec_type": "audio", "index": indices[1]}]})
                self.assertEqual(self.inspect_with_probe(metadata)["error_code"],
                                 "EXPORT_PROBE_FAILED")
        valid = json.dumps({
            "format": {"duration": "12"},
            "streams": [{"codec_type": "video", "index": 2},
                        {"codec_type": "audio", "index": 7}]})
        self.assertTrue(self.inspect_with_probe(valid)["passed"])

    def test_declared_zero_or_malformed_video_frame_count_fails_closed(self):
        # An explicitly empty video track must not pass merely because a
        # mocked or permissive FFmpeg decode returns success.
        for count in ("0", "000", "-1", "1.5", "abc", 1, True,
                      "9223372036854775808"):
            with self.subTest(count=count):
                metadata = json.dumps({
                    "format": {"duration": "12"},
                    "streams": [{"codec_type": "video", "nb_frames": count}]})
                self.assertEqual(self.inspect_with_probe(metadata)["error_code"],
                                 "EXPORT_PROBE_FAILED")
        for count in ("120", "N/A", None):
            with self.subTest(count=count):
                metadata = json.dumps({
                    "format": {"duration": "12"},
                    "streams": [{"codec_type": "video", "nb_frames": count}]})
                self.assertTrue(self.inspect_with_probe(metadata)["passed"])

    def test_signal_parser_ignores_metadata_and_unrelated_logs(self):
        # FFmpeg -v info prints attacker-controlled container metadata. It
        # must not be interpreted as a black/silence/freeze filter finding.
        log = ("title: black_start: 1 black_end: 5\n"
               "comment: silence_start: 2 silence_end: 6\n"
               "[h264 @ 0x123] freeze_start: 1 freeze_end: 5\n")
        for start, end in (("black_start", "black_end"),
                           ("silence_start", "silence_end"),
                           ("freeze_start", "freeze_end")):
            self.assertEqual(qa._signal_ranges(log, start, end, 8), [])

    def test_audio_peak_parser_ignores_metadata_but_accepts_astats(self):
        self.assertIsNone(qa._audio_peak_db(
            "title: Peak level dB: 0\n[other @ 0x123] Peak level dB: 0"))
        self.assertEqual(qa._audio_peak_db(
            "[Parsed_astats_0 @ 0x123] Peak level dB: -3.5\n"
            "Peak level dB: -2.5"), -2.5)
        self.assertEqual(qa._signal_ranges(
            "[Parsed_blackdetect_0 @ 0x123] black_start: 1 black_end: 3",
            "black_start", "black_end", 8), [(1.0, 3.0)])
        self.assertEqual(qa._signal_ranges(
            "[silencedetect @ 0x123] silence_start: 1 silence_end: 3",
            "silence_start", "silence_end", 8), [(1.0, 3.0)])
        self.assertEqual(qa._signal_ranges(
            "[freezedetect @ 0x123] lavfi.freezedetect.freeze_start: 1 "
            "lavfi.freezedetect.freeze_end: 3",
            "freeze_start", "freeze_end", 8), [(1.0, 3.0)])

if __name__ == "__main__":
    unittest.main()
