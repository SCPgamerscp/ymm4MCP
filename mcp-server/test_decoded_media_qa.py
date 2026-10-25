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
        self.assertEqual(len(result["samples"]), 2)
        self.assertEqual(len(commands), 5)
        self.assertEqual([c[0][0] for c in commands], ["ffprobe", "ffmpeg", "ffmpeg", "ffmpeg", "ffmpeg"])
        self.assertTrue(all(str(self.video) in cmd for cmd, _ in commands))
        self.assertTrue(all(isinstance(cmd, list) for cmd, _ in commands))
        self.assertTrue(all(timeout <= 30 for _, timeout in commands))

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
        self.assertEqual(len(report["black_frames"]), 2)
        self.assertEqual(len(report["long_silences"]), 2)
        self.assertEqual(report["long_silences"][0]["duration_seconds"], 3.5)
        self.assertEqual(report["black_frames"][1]["sample_start_seconds"], 22.5)
        self.assertEqual({item["severity"] for item in report["issues"]}, {"warning"})
        self.assertEqual({item["code"] for item in report["issues"]},
                         {"EXPORT_SAMPLE_BLACK_FRAMES", "EXPORT_SAMPLE_LONG_SILENCE"})

    def test_incomplete_signal_events_do_not_create_false_reports(self):
        self.assertEqual(qa._signal_ranges("silence_start: 2", "silence_start", "silence_end"), [])
        self.assertEqual(qa._signal_ranges("silence_end: 4", "silence_start", "silence_end"), [])
        self.assertEqual(qa._signal_ranges("black_start: 5 black_end: 4",
                                           "black_start", "black_end"), [])

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
                       {"sample_seconds": 16}, {"require_audio": "false"}):
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
                                            "require_audio": False})
        self.assertEqual(actual, report)
        inspect.assert_called_once_with("/tmp/clip.mp4", require_audio=False, sample_seconds=3)


if __name__ == "__main__":
    unittest.main()
