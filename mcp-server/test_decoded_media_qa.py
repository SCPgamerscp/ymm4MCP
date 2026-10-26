"""Regression tests for sampled decode QA; no FFmpeg or YMM4 required."""
import json
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import decoded_media_qa as qa
import server


def process(returncode=0, stdout="", stderr=""):
    return subprocess.CompletedProcess([], returncode, stdout, stderr)


class DecodedMediaQATests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.video = Path(self.temp.name) / "video name.mp4"
        self.video.write_bytes(b"not a real media file" + b"0" * 32)
        self.metadata = json.dumps({
            "format": {"duration": "30.5"},
            "streams": [{"codec_type": "video", "width": 1920, "height": 1080},
                        {"codec_type": "audio", "channels": 2}],
        })

    def test_both_ends_and_both_streams_are_decoded(self):
        commands = []
        def fake(command, timeout):
            commands.append((command, timeout))
            return process(stdout=self.metadata if command[0] == "ffprobe" else "")
        with patch.object(qa, "run", side_effect=fake):
            result = qa.inspect(str(self.video))
        self.assertTrue(result["passed"])
        self.assertEqual(result["verification_scope"], "sampled_audio_video_decode")
        self.assertFalse(result["fully_decoded"])
        self.assertEqual(result["video_streams"], 1)
        self.assertEqual(result["audio_streams"], 1)
        self.assertEqual(len(result["samples"]), 3)
        self.assertEqual(len(commands), 7)
        self.assertEqual([c[0][0] for c in commands], ["ffprobe"] + ["ffmpeg"] * 6)
        self.assertTrue(all(str(self.video) in cmd for cmd, _ in commands))
        self.assertTrue(all(isinstance(cmd, list) for cmd, _ in commands))
        self.assertTrue(all(timeout <= 30 for _, timeout in commands))
        self.assertTrue(all("-abort_on" in cmd and
                            cmd[cmd.index("-abort_on") + 1] == "empty_output"
                            for cmd, _ in commands if cmd[0] == "ffmpeg"))

    def test_full_decode_checks_all_streams_and_marks_full_scope(self):
        commands = []
        def fake(command, timeout):
            commands.append((command, timeout))
            return process(stdout=self.metadata if command[0] == "ffprobe" else "")
        with patch.object(qa, "run", side_effect=fake):
            result = qa.inspect(str(self.video), full_decode=True)
        self.assertTrue(result["passed"], result)
        self.assertTrue(result["fully_decoded"])
        self.assertEqual(result["verification_scope"], "full_audio_video_decode")
        self.assertEqual(result["signal_scope"], "sampled_only")
        self.assertEqual(len(commands), 9)  # probe, six samples, two full streams
        full_commands = [cmd for cmd, _ in commands if cmd[0] == "ffmpeg"
                         and "-v" in cmd and cmd[cmd.index("-v") + 1] == "error"]
        self.assertEqual(len(full_commands), 2)
        self.assertEqual({cmd[cmd.index("-map") + 1] for cmd in full_commands},
                         {"0:v", "0:a"})
        self.assertTrue(all("-ss" not in cmd and "-t" not in cmd
                            and "-abort_on" in cmd for cmd in full_commands))
        self.assertTrue(all(cmd[cmd.index("-abort_on") + 1]
                            == "empty_output+empty_output_stream" for cmd in full_commands))
        sampled_commands = [cmd for cmd, _ in commands if cmd[0] == "ffmpeg"
                            and "-ss" in cmd]
        self.assertTrue(all(cmd[cmd.index("-abort_on") + 1] == "empty_output"
                            for cmd in sampled_commands))
        self.assertTrue(all(timeout <= 900 for _, timeout in commands))

    def test_full_decode_fails_if_second_audio_stream_is_empty(self):
        # With only empty_output, FFmpeg can succeed because the first mapped
        # audio stream produced packets even though the second produced none.
        metadata = json.dumps({
            "format": {"duration": "30.5"},
            "streams": [{"codec_type": "video"}, {"codec_type": "audio"},
                        {"codec_type": "audio"}],
        })
        def fake(command, timeout):
            if command[0] == "ffprobe":
                return process(stdout=metadata)
            if "-map" in command and command[command.index("-map") + 1] == "0:a":
                abort_flags = command[command.index("-abort_on") + 1]
                if "empty_output_stream" not in abort_flags:
                    return process()  # Reproduces the old false-positive path.
                return process(1, stderr="second mapped audio stream is empty")
            return process()
        with patch.object(qa, "run", side_effect=fake):
            report = qa.inspect(str(self.video), full_decode=True)
        self.assertEqual(report["audio_streams"], 2)
        self.assertFalse(report["passed"])
        self.assertFalse(report["fully_decoded"])
        self.assertIn("EXPORT_FULL_DECODE_FAILED",
                      {issue["code"] for issue in report["issues"]})

    def test_full_decode_detects_middle_corruption_not_seen_in_samples(self):
        def fake(command, timeout):
            if command[0] == "ffprobe":
                return process(stdout=self.metadata)
            if "-v" in command and command[command.index("-v") + 1] == "error":
                if "0:v" in command:
                    return process(1, stderr="corrupt packet in middle")
            return process()
        with patch.object(qa, "run", side_effect=fake):
            sampled = qa.inspect(str(self.video))
            full = qa.inspect(str(self.video), full_decode=True)
        self.assertTrue(sampled["passed"])
        self.assertFalse(sampled["fully_decoded"])
        self.assertFalse(full["passed"])
        self.assertFalse(full["fully_decoded"])
        self.assertIn("EXPORT_FULL_DECODE_FAILED",
                      {issue["code"] for issue in full["issues"]})

    def test_full_decode_timeout_and_duration_limit_fail_closed(self):
        def timeout(command, seconds):
            if command[0] == "ffprobe":
                return process(stdout=self.metadata)
            if "-v" in command and command[command.index("-v") + 1] == "error":
                raise subprocess.TimeoutExpired(command, seconds)
            return process()
        with patch.object(qa, "run", side_effect=timeout):
            report = qa.inspect(str(self.video), full_decode=True)
        self.assertFalse(report["passed"])
        self.assertIn("EXPORT_FULL_DECODE_TIMEOUT",
                      {issue["code"] for issue in report["issues"]})
        too_long = json.dumps({"format": {"duration": "901"}, "streams":
                               [{"codec_type": "video"}]})
        with patch.object(qa, "run", return_value=process(stdout=too_long)) as run:
            report = qa.inspect(str(self.video), full_decode=True)
        self.assertEqual(report["error_code"], "EXPORT_FULL_DECODE_DURATION_LIMIT")
        self.assertEqual(run.call_count, 1)

    def test_sampled_black_frames_and_long_silence_are_nonblocking_warnings(self):
        def fake(command, timeout):
            if command[0] == "ffprobe":
                return process(stdout=self.metadata)
            if "-vf" in command:
                return process(stderr="[blackdetect @ 0x1] black_start:0 black_end:2 black_duration:2")
            return process(stderr=("[silencedetect @ 0x1] silence_start: 1\n"
                                   "[silencedetect @ 0x1] silence_end: 4.5 | silence_duration: 3.5"))
        with patch.object(qa, "run", side_effect=fake):
            report = qa.inspect(str(self.video))
        self.assertTrue(report["passed"])
        self.assertEqual(report["signal_scope"], "sampled_only")
        self.assertEqual(len(report["black_frames"]), 3)
        self.assertEqual(len(report["long_silences"]), 3)
        self.assertEqual(report["long_silences"][0]["duration_seconds"], 3.5)
        self.assertEqual(report["black_frames"][2]["sample_start_seconds"], 22.5)
        self.assertEqual({item["severity"] for item in report["issues"]}, {"warning"})
        self.assertEqual({item["code"] for item in report["issues"]},
                         {"EXPORT_SAMPLE_BLACK_FRAMES", "EXPORT_SAMPLE_LONG_SILENCE"})

    def test_incomplete_signal_events_do_not_create_false_reports(self):
        self.assertEqual(qa._signal_ranges("silence_start: 2", "silence_start", "silence_end"), [])
        self.assertEqual(qa._signal_ranges("silence_end: 4", "silence_start", "silence_end"), [])
        self.assertEqual(qa._signal_ranges("black_start: 5 black_end: 4",
                                           "black_start", "black_end"), [])

    def test_sampled_audio_peak_near_clipping_is_warning_not_proof(self):
        self.assertEqual(qa._audio_peak_db("Peak level dB: -3.25\nPeak level dB: -1.5"), -1.5)
        self.assertIsNone(qa._audio_peak_db("Peak level dB: -inf"))
        def fake(command, timeout):
            if command[0] == "ffprobe":
                return process(stdout=self.metadata)
            if "-af" in command:
                return process(stderr="[Parsed_astats @ 0x1] Peak level dB: 0.000000")
            return process()
        with patch.object(qa, "run", side_effect=fake):
            report = qa.inspect(str(self.video))
        self.assertTrue(report["success"])
        self.assertTrue(report["passed"])  # near-clipping is not proof of distortion
        self.assertEqual(len(report["audio_peaks"]), 3)
        self.assertEqual(report["audio_peaks"][0]["peak_dbfs"], 0.0)
        self.assertEqual([i["code"] for i in report["issues"]],
                         ["EXPORT_SAMPLE_AUDIO_PEAK_NEAR_CLIPPING"] * 3)

    def test_normal_audio_peak_has_no_clipping_warning(self):
        def fake(command, timeout):
            if command[0] == "ffprobe":
                return process(stdout=self.metadata)
            if "-af" in command:
                return process(stderr="Peak level dB: -6.020600")
            return process()
        with patch.object(qa, "run", side_effect=fake):
            report = qa.inspect(str(self.video))
        self.assertTrue(report["passed"])
        self.assertEqual(len(report["audio_peaks"]), 3)
        self.assertEqual(report["issues"], [])

    def test_decode_failure_does_not_claim_success(self):
        def fake(command, timeout):
            if command[0] == "ffprobe":
                return process(stdout=self.metadata)
            if "0:a:0" in command:
                return process(1, stderr="corrupt audio packet")
            return process()
        with patch.object(qa, "run", side_effect=fake):
            result = qa.inspect(str(self.video))
        self.assertTrue(result["success"])
        self.assertFalse(result["passed"])
        self.assertEqual(result["issues"][0]["code"], "EXPORT_DECODE_FAILED")
        self.assertEqual(result["issues"][0]["stream_type"], "audio")

    def test_missing_audio_and_video_stream(self):
        data = json.dumps({"format": {"duration": "2.0"}, "streams": []})
        with patch.object(qa, "run", return_value=process(stdout=data)):
            report = qa.inspect(str(self.video))
        self.assertFalse(report["passed"])
        self.assertEqual({i["code"] for i in report["issues"]},
                         {"EXPORT_VIDEO_STREAM_MISSING", "EXPORT_AUDIO_STREAM_MISSING"})
        with patch.object(qa, "run", return_value=process(stdout=data)):
            report = qa.inspect(str(self.video), require_audio=False)
        self.assertEqual([i["code"] for i in report["issues"]], ["EXPORT_VIDEO_STREAM_MISSING"])

    def test_invalid_arguments_and_missing_files(self):
        for kwargs in ({"sample_seconds": True}, {"sample_seconds": 0},
                       {"sample_seconds": 16}, {"require_audio": "false"},
                       {"full_decode": "true"}, {"full_decode": 1}):
            with self.subTest(kwargs=kwargs), self.assertRaises(ValueError):
                qa.inspect(str(self.video), **kwargs)
        with self.assertRaises(ValueError):
            qa.inspect("relative.mp4")
        self.assertEqual(qa.inspect(str(Path(self.temp.name) / "missing.mp4"))["error_code"],
                         "EXPORT_FILE_MISSING")

    def test_probe_errors_fail_closed(self):
        with patch.object(qa, "run", side_effect=FileNotFoundError):
            self.assertEqual(qa.inspect(str(self.video))["error_code"], "FFMPEG_UNAVAILABLE")
        with patch.object(qa, "run", side_effect=subprocess.TimeoutExpired("ffprobe", 15)):
            self.assertEqual(qa.inspect(str(self.video))["error_code"], "EXPORT_PROBE_TIMEOUT")
        with patch.object(qa, "run", return_value=process(stdout="{bad")):
            self.assertEqual(qa.inspect(str(self.video))["error_code"], "EXPORT_PROBE_FAILED")


class McpDispatchTests(unittest.IsolatedAsyncioTestCase):
    async def test_get_info_calls_decode_inspector(self):
        report = {"success": True, "passed": True, "issues": []}
        with patch.object(server.decoded_media_qa, "inspect", return_value=report) as inspect:
            actual = await server.dispatch({"action": "get_info", "sub_action": "decoded_export_qa",
                                            "path": "/tmp/clip.mp4", "sample_seconds": 3,
                                            "require_audio": False, "full_decode": True})
        self.assertEqual(actual, report)
        inspect.assert_called_once_with("/tmp/clip.mp4", require_audio=False,
                                        sample_seconds=3, full_decode=True)


if __name__ == "__main__":
    unittest.main()
