"""Full-track integrated LUFS quality-gate regression tests (Issue #61)."""
import json
import shutil
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import decoded_media_qa as qa

LOUDNORM = """{
    "input_i" : "-20.50",
    "input_tp" : "-4.10",
    "input_lra" : "0.00",
    "input_thresh" : "-30.50",
    "output_i" : "-16.00"
}"""


def completed(stdout="", stderr="", returncode=0):
    return subprocess.CompletedProcess([], returncode, stdout, stderr)


class LoudnessTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.video = Path(self.temp.name) / "example.mp4"
        self.video.write_bytes(b"fake media" + b"x" * 64)
        self.metadata = json.dumps({
            "format": {"duration": "3"},
            "streams": [{"codec_type": "video", "width": 160, "height": 90},
                        {"codec_type": "audio", "channels": 2}],
        })
        self.commands = []

    def fake_run(self, command, timeout):
        self.commands.append(command)
        if command[0] == "ffprobe":
            return completed(stdout=self.metadata)
        if "-af" in command and "loudnorm=" in command[command.index("-af") + 1]:
            return completed(stderr=LOUDNORM)
        return completed()

    def test_loudnorm_parser_rejects_missing_duplicate_and_nonfinite(self):
        self.assertEqual(-20.5, qa._integrated_lufs_from_loudnorm(LOUDNORM))
        for diagnostic in ("", LOUDNORM + "\n" + LOUDNORM,
                           LOUDNORM.replace("-20.50", "-inf"),
                           LOUDNORM.replace('"input_tp"', '"missing_tp"'),
                           LOUDNORM.replace('"input_i" : "-20.50"',
                                            '"input_i" : "-20.50", "input_i" : "-15.00"')):
            with self.subTest(diagnostic=diagnostic[:30]):
                with self.assertRaises(ValueError):
                    qa._integrated_lufs_from_loudnorm(diagnostic)

    def test_loudness_gates_are_opt_in_and_require_full_decode(self):
        with self.assertRaises(ValueError):
            qa.inspect(str(self.video), min_integrated_lufs=-30)
        with self.assertRaises(ValueError):
            qa.inspect(str(self.video), full_decode=True, min_integrated_lufs=-10,
                       max_integrated_lufs=-30)
        with self.assertRaises(ValueError):
            qa.inspect(str(self.video), full_decode=True, min_integrated_lufs=True)
        with patch.object(qa, "run", side_effect=self.fake_run):
            baseline = qa.inspect(str(self.video), full_decode=True)
            passed = qa.inspect(str(self.video), full_decode=True,
                                min_integrated_lufs=-25, max_integrated_lufs=-15)
            quiet = qa.inspect(str(self.video), full_decode=True,
                               min_integrated_lufs=-18)
            loud = qa.inspect(str(self.video), full_decode=True,
                              max_integrated_lufs=-22)
        self.assertTrue(baseline["passed"], baseline)
        self.assertEqual([], baseline["integrated_loudness"])
        self.assertTrue(passed["passed"], passed)
        self.assertEqual([{"stream_ordinal": 0, "integrated_lufs": -20.5}],
                         passed["integrated_loudness"])
        self.assertFalse(quiet["passed"])
        self.assertIn("EXPORT_LOUDNESS_TOO_QUIET",
                      {issue["code"] for issue in quiet["issues"]})
        self.assertFalse(loud["passed"])
        self.assertIn("EXPORT_LOUDNESS_TOO_LOUD",
                      {issue["code"] for issue in loud["issues"]})
        loudnorm_commands = [command for command in self.commands
                             if "-af" in command and "loudnorm=" in command[command.index("-af") + 1]]
        self.assertEqual(3, len(loudnorm_commands))
        self.assertTrue(all(command[command.index("-map") + 1] == "0:a:0"
                            for command in loudnorm_commands))

    def test_invalid_measurement_and_process_failure_are_blocking(self):
        def missing_measurement(command, timeout):
            result = self.fake_run(command, timeout)
            if "-af" in command and "loudnorm=" in command[command.index("-af") + 1]:
                return completed(stderr="filter did not report a measurement")
            return result
        with patch.object(qa, "run", side_effect=missing_measurement):
            report = qa.inspect(str(self.video), full_decode=True,
                                min_integrated_lufs=-30)
        self.assertFalse(report["passed"])
        self.assertIn("EXPORT_LOUDNESS_UNVERIFIABLE",
                      {issue["code"] for issue in report["issues"]})


@unittest.skipUnless(shutil.which("ffmpeg") and shutil.which("ffprobe"),
                     "FFmpeg is required for the real loudness integration test")
class RealLoudnessTests(unittest.TestCase):
    def test_sine_audio_is_measured_over_complete_track(self):
        with tempfile.TemporaryDirectory() as directory:
            video = Path(directory) / "sine.mp4"
            subprocess.run(
                ["ffmpeg", "-hide_banner", "-nostdin", "-loglevel", "error",
                 "-f", "lavfi", "-i", "color=c=blue:s=160x90:r=10:d=3",
                 "-f", "lavfi", "-i", "sine=frequency=1000:duration=3",
                 "-c:v", "mpeg4", "-c:a", "aac", "-shortest", "-y", str(video)],
                check=True, capture_output=True, timeout=30)
            report = qa.inspect(str(video), full_decode=True,
                                min_integrated_lufs=-70, max_integrated_lufs=10)
            self.assertTrue(report["passed"], report)
            self.assertEqual(1, len(report["integrated_loudness"]))
            measured = report["integrated_loudness"][0]["integrated_lufs"]
            self.assertTrue(-50 < measured < 0, report)
            strict = qa.inspect(str(video), full_decode=True,
                                min_integrated_lufs=0)
            self.assertFalse(strict["passed"], strict)
            self.assertIn("EXPORT_LOUDNESS_TOO_QUIET",
                          {issue["code"] for issue in strict["issues"]})


if __name__ == "__main__":
    unittest.main()
