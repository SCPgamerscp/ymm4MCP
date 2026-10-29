"""Regression tests for incomplete FFmpeg audio diagnostics (Issue #61)."""
import json
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import decoded_media_qa as qa


VALID_LOUDNORM = """{
    "input_i": "-20.00",
    "input_tp": "-4.00",
    "input_lra": "0.00",
    "input_thresh": "-30.00",
    "output_i": "-16.00"
}"""
BAD_JSON = """{
    "input_i": "-21.00",
}"""
DUPLICATE_KEY = """{
    "input_i": "-21.00",
    "input_i": "-15.00"
}"""
MISSING_INTEGRATED = """{
    "input_tp": "-4.00"
}"""


def completed(stdout="", stderr=""):
    return subprocess.CompletedProcess([], 0, stdout, stderr)


class AudioDiagnosticIntegrityTests(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.video = Path(temp.name) / "clip.mp4"
        self.video.write_bytes(b"test media" + b"x" * 64)
        self.metadata = json.dumps({
            "format": {"duration": "3"},
            "streams": [{"codec_type": "video", "width": 160, "height": 90},
                        {"codec_type": "audio", "channels": 2}],
        })

    def test_unmeasured_second_overall_cannot_certify_silence(self):
        self.assertTrue(qa._audio_overall_silent(
            "Overall\nPeak level dB: -inf"))
        self.assertTrue(qa._audio_overall_silent(
            "Overall\nPeak level dB: -inf\nOverall\nPeak level dB: -inf"))
        self.assertFalse(qa._audio_overall_silent(
            "Overall\nPeak level dB: -inf\nOverall"))
        self.assertFalse(qa._audio_overall_silent(
            "Overall\nPeak level dB: -8\nOverall\nPeak level dB: -inf"))

    def test_strict_peak_gate_rejects_incomplete_last_overall(self):
        def fake_run(command, timeout):
            if command[0] == "ffprobe":
                return completed(stdout=self.metadata)
            if "-af" in command:
                return completed(stderr="Overall\nPeak level dB: -inf\nOverall")
            return completed()

        with patch.object(qa, "run", side_effect=fake_run):
            report = qa.inspect(str(self.video), require_audio_peak_metrics=True)
        self.assertFalse(report["passed"], report)
        codes = {issue["code"] for issue in report["issues"]}
        self.assertIn("EXPORT_SAMPLE_AUDIO_METRICS_INCOMPLETE", codes)

    def test_valid_and_unrelated_json_remain_accepted(self):
        self.assertEqual((-20.0, -4.0),
                         qa._loudnorm_measurements(VALID_LOUDNORM))
        unrelated = '{\n    "filter": "metadata"\n}'
        self.assertEqual((-20.0, -4.0),
                         qa._loudnorm_measurements(unrelated + "\n" + VALID_LOUDNORM))

    def test_invalid_secondary_loudnorm_measurements_fail_closed(self):
        for invalid in (BAD_JSON, DUPLICATE_KEY, MISSING_INTEGRATED):
            for order in ((VALID_LOUDNORM, invalid), (invalid, VALID_LOUDNORM)):
                with self.subTest(invalid=invalid, order=order[0][:20]):
                    with self.assertRaises(ValueError):
                        qa._loudnorm_measurements("\n".join(order))

    def test_full_decode_loudness_gate_rejects_corrupt_secondary_block(self):
        def fake_run(command, timeout):
            if command[0] == "ffprobe":
                return completed(stdout=self.metadata)
            if "-af" in command and "loudnorm=" in command[command.index("-af") + 1]:
                return completed(stderr=VALID_LOUDNORM + "\n" + BAD_JSON)
            return completed()

        with patch.object(qa, "run", side_effect=fake_run):
            report = qa.inspect(str(self.video), full_decode=True,
                                min_integrated_lufs=-30)
        self.assertFalse(report["passed"], report)
        self.assertIn("EXPORT_LOUDNESS_UNVERIFIABLE",
                      {issue["code"] for issue in report["issues"]})


if __name__ == "__main__":
    unittest.main()
