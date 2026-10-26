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


if __name__ == "__main__":
    unittest.main()
