"""Regression tests for FFmpeg astats peak and silence consistency (#61)."""
import json
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import decoded_media_qa as qa


class AstatsConsistencyTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.video = Path(self.directory.name) / "sample.mp4"
        self.video.write_bytes(b"test media" * 8)

    def inspect_audio(self, log, *, channels=None):
        audio = {"codec_type": "audio"}
        if channels is not None:
            audio["channels"] = channels
        metadata = json.dumps({
            "format": {"duration": "1.5"},
            "streams": [{"codec_type": "video"}, audio],
        })

        def fake(command, timeout):
            if command[0] == "ffprobe":
                return subprocess.CompletedProcess(command, 0, metadata, "")
            return subprocess.CompletedProcess(
                command, 0, "", log if "0:a:0" in command else "")

        with patch.object(qa, "run", side_effect=fake):
            return qa.inspect(str(self.video))

    def test_repeated_channel_sections_keep_loudest_peak(self):
        log = ("[Parsed_astats_0 @ 0x123] Channel: 1\n"
               "[Parsed_astats_0 @ 0x123] Peak level dB: -3\n"
               "[Parsed_astats_0 @ 0x123] Channel: 2\n"
               "[Parsed_astats_0 @ 0x123] Peak level dB: -30\n"
               "[Parsed_astats_0 @ 0x123] Channel: 1\n"
               "[Parsed_astats_0 @ 0x123] Peak level dB: -inf\n"
               "[Parsed_astats_0 @ 0x123] Channel: 2\n"
               "[Parsed_astats_0 @ 0x123] Peak level dB: -90\n")
        self.assertEqual(qa._audio_channel_peaks(log), {1: -3.0, 2: -30.0})
        result = self.inspect_audio(log, channels=2)
        self.assertTrue(result["passed"], result)
        self.assertEqual(result["audio_channel_peaks"][0]["peak_dbfs_by_channel"],
                         {"1": -3.0, "2": -30.0})
        self.assertIn("EXPORT_SAMPLE_AUDIO_STEREO_IMBALANCE",
                      [issue["code"] for issue in result["issues"]])
        self.assertNotIn("EXPORT_SAMPLE_AUDIO_ALL_CHANNELS_SILENT",
                         [issue["code"] for issue in result["issues"]])

    def test_repeated_overall_summaries_do_not_hide_audible_audio(self):
        log = ("[Parsed_astats_0 @ 0x123] Overall\n"
               "[Parsed_astats_0 @ 0x123] Peak level dB: -8\n"
               "[Parsed_astats_0 @ 0x123] Overall\n"
               "[Parsed_astats_0 @ 0x123] Peak level dB: -inf\n")
        self.assertFalse(qa._audio_overall_silent(log))
        result = self.inspect_audio(log)
        self.assertTrue(result["passed"], result)
        self.assertNotIn("EXPORT_SAMPLE_AUDIO_ALL_CHANNELS_SILENT",
                         [issue["code"] for issue in result["issues"]])

    def test_overall_silence_counts_when_channel_count_is_unknown(self):
        log = ("[Parsed_astats_0 @ 0x123] Channel: 1\n"
               "[Parsed_astats_0 @ 0x123] Peak level dB: -inf\n"
               "[Parsed_astats_0 @ 0x123] Overall\n"
               "[Parsed_astats_0 @ 0x123] Peak level dB: -inf\n")
        result = self.inspect_audio(log)
        self.assertTrue(result["passed"], result)
        self.assertEqual(
            sum(i["code"] == "EXPORT_SAMPLE_AUDIO_ALL_CHANNELS_SILENT"
                for i in result["issues"]), 1, result)

    def test_overall_silence_counts_even_if_channel_sections_are_incomplete(self):
        log = ("[Parsed_astats_0 @ 0x123] Channel: 1\n"
               "[Parsed_astats_0 @ 0x123] Peak level dB: -inf\n"
               "[Parsed_astats_0 @ 0x123] Overall\n"
               "[Parsed_astats_0 @ 0x123] Peak level dB: -inf\n")
        result = self.inspect_audio(log, channels=2)
        self.assertTrue(result["passed"], result)
        codes = [i["code"] for i in result["issues"]]
        self.assertIn("EXPORT_SAMPLE_AUDIO_CHANNEL_COUNT_MISMATCH", codes)
        self.assertEqual(codes.count("EXPORT_SAMPLE_AUDIO_ALL_CHANNELS_SILENT"), 1)

    def test_partial_channel_sections_without_overall_do_not_claim_full_silence(self):
        log = ("[Parsed_astats_0 @ 0x123] Channel: 1\n"
               "[Parsed_astats_0 @ 0x123] Peak level dB: -inf\n")
        result = self.inspect_audio(log, channels=2)
        self.assertTrue(result["passed"], result)
        codes = [i["code"] for i in result["issues"]]
        self.assertIn("EXPORT_SAMPLE_AUDIO_CHANNEL_COUNT_MISMATCH", codes)
        self.assertNotIn("EXPORT_SAMPLE_AUDIO_ALL_CHANNELS_SILENT", codes)


    def test_oversized_channel_id_is_ignored_without_integer_conversion(self):
        # Thousands of digits must not raise Python's int conversion limit.
        log = ("[Parsed_astats_0 @ 0x123] Channel: " + "9" * 5000 + "\n"
               "[Parsed_astats_0 @ 0x123] Peak level dB: -2\n")
        self.assertEqual(qa._audio_channel_peaks(log), {})
        result = self.inspect_audio(log, channels=2)
        # Invalid channel headers cannot be trusted as successful audio QA,
        # even though the parser safely avoids converting huge integers.
        self.assertFalse(result["passed"], result)
        self.assertIn("EXPORT_SAMPLE_AUDIO_CHANNEL_HEADER_INVALID",
                      [issue["code"] for issue in result["issues"]])

    def test_invalid_channel_header_clears_previous_channel(self):
        log = ("[Parsed_astats_0 @ 0x123] Channel: 1\n"
               "[Parsed_astats_0 @ 0x123] Peak level dB: -18\n"
               "[Parsed_astats_0 @ 0x123] Channel: 65\n"
               "[Parsed_astats_0 @ 0x123] Peak level dB: -1\n")
        self.assertEqual(qa._audio_channel_peaks(log), {1: -18.0})


if __name__ == "__main__":
    unittest.main()
