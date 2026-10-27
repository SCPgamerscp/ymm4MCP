"""Regression tests for required-track tail completeness in completed export QA (#61)."""
import json
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import decoded_media_qa as qa


class RequiredTrackTailTests(unittest.TestCase):
    def setUp(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.video = Path(directory.name) / "completed.mp4"
        self.video.write_bytes(b"synthetic export" * 8)

    def inspect(self, *, video_duration="4.0", audio_durations=("1.0",),
                require_audio=True, **options):
        streams = [
            {"index": 0, "codec_type": "video", "start_time": "0",
             "duration": video_duration, "width": 160, "height": 90,
             "avg_frame_rate": "10/1"},
        ]
        for ordinal, duration in enumerate(audio_durations, 1):
            streams.append({
                "index": ordinal, "codec_type": "audio", "start_time": "0",
                "duration": duration, "sample_rate": "44100", "channels": 1,
            })
        metadata = json.dumps({
            "format": {"duration": "4.0", "start_time": "0"},
            "streams": streams,
        })

        def fake_run(command, timeout):
            if command[0] == "ffprobe":
                return subprocess.CompletedProcess(command, 0, metadata, "")
            return subprocess.CompletedProcess(command, 0, "", "")

        with patch.object(qa, "run", side_effect=fake_run):
            return qa.inspect(str(self.video), sample_seconds=1,
                              require_audio=require_audio, **options)

    def tail(self, result, kind, ordinal=0):
        matches = [issue for issue in result["issues"]
                   if issue["code"] == "EXPORT_STREAM_ENDS_EARLY"
                   and issue["stream_type"] == kind
                   and issue["stream_ordinal"] == ordinal]
        self.assertEqual(len(matches), 1, result)
        return matches[0]

    def test_required_primary_audio_ending_early_fails_qa(self):
        result = self.inspect()
        self.assertTrue(result["success"], result)
        self.assertFalse(result["passed"], result)
        self.assertEqual(self.tail(result, "audio")["severity"], "error")

    def test_optional_audio_ending_early_is_advisory(self):
        result = self.inspect(require_audio=False)
        self.assertTrue(result["passed"], result)
        self.assertEqual(self.tail(result, "audio")["severity"], "warning")

    def test_expected_audio_format_makes_primary_audio_required(self):
        result = self.inspect(require_audio=False, expected_audio_sample_rate=44100)
        self.assertFalse(result["passed"], result)
        self.assertEqual(self.tail(result, "audio")["severity"], "error")

    def test_primary_video_ending_early_fails_qa(self):
        result = self.inspect(video_duration="1.0", audio_durations=("4.0",))
        self.assertFalse(result["passed"], result)
        self.assertEqual(self.tail(result, "video")["severity"], "error")

    def test_short_secondary_audio_remains_advisory(self):
        result = self.inspect(audio_durations=("4.0", "1.0"))
        self.assertTrue(result["passed"], result)
        self.assertEqual(self.tail(result, "audio", 1)["severity"], "warning")


if __name__ == "__main__":
    unittest.main()
