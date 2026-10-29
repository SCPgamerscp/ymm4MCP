"""Optional real-FFmpeg smoke test for sampled visual/audio QA."""
import shutil
import subprocess
import tempfile
import unittest
from pathlib import Path

import decoded_media_qa


@unittest.skipUnless(shutil.which("ffmpeg") and shutil.which("ffprobe"),
                     "ffmpeg and ffprobe must be on PATH")
class DecodedMediaFFmpegIntegrationTests(unittest.TestCase):
    def test_black_video_and_silent_audio_are_reported_as_sampled_warnings(self):
        with tempfile.TemporaryDirectory() as directory:
            video = Path(directory) / "sample.mp4"
            subprocess.run(
                ["ffmpeg", "-hide_banner", "-nostdin", "-loglevel", "error",
                 "-f", "lavfi", "-i", "color=c=black:s=160x90:r=10:d=4",
                 "-f", "lavfi", "-i", "anullsrc=r=44100:cl=stereo",
                 "-t", "4", "-c:v", "mpeg4", "-c:a", "aac", "-y", str(video)],
                check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                timeout=30)
            result = decoded_media_qa.inspect(str(video), sample_seconds=4, full_decode=True)
            self.assertTrue(result["success"], result)
            self.assertTrue(result["passed"], result)  # warnings are not errors
            self.assertTrue(result["fully_decoded"], result)
            self.assertEqual(result["verification_scope"], "full_audio_video_decode")
            self.assertEqual(result["signal_scope"], "sampled_only")
            self.assertGreaterEqual(len(result["black_frames"]), 1)
            self.assertGreaterEqual(len(result["long_silences"]), 1)
            self.assertTrue(all(issue["severity"] == "warning"
                                for issue in result["issues"]))


    def test_loud_audio_reports_near_clipping_sample_warning(self):
        with tempfile.TemporaryDirectory() as directory:
            video = Path(directory) / "loud.mp4"
            subprocess.run(
                ["ffmpeg", "-hide_banner", "-nostdin", "-loglevel", "error",
                 "-f", "lavfi", "-i", "color=c=blue:s=160x90:r=10:d=3",
                 "-f", "lavfi", "-i", "sine=frequency=1000:duration=3",
                 "-af", "volume=24dB", "-t", "3",
                 "-c:v", "mpeg4", "-c:a", "aac", "-y", str(video)],
                check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                timeout=30)
            result = decoded_media_qa.inspect(str(video), sample_seconds=3)
            self.assertTrue(result["success"], result)
            self.assertTrue(result["passed"], result)
            self.assertEqual(len(result["audio_peaks"]), 1)
            self.assertGreaterEqual(result["audio_peaks"][0]["peak_dbfs"], -0.1)
            self.assertIn("EXPORT_SAMPLE_AUDIO_PEAK_NEAR_CLIPPING",
                          {issue["code"] for issue in result["issues"]})


    def test_short_audio_stream_is_advisory_unless_strict_tail_required(self):
        """Real FFmpeg metadata detects the gap; strict policy determines pass/fail."""
        with tempfile.TemporaryDirectory() as directory:
            video = Path(directory) / "short_audio.mp4"
            subprocess.run(
                ["ffmpeg", "-hide_banner", "-nostdin", "-loglevel", "error",
                 "-f", "lavfi", "-i", "color=c=blue:s=160x90:r=10:d=4",
                 "-f", "lavfi", "-i", "sine=frequency=440:duration=1",
                 "-map", "0:v:0", "-map", "1:a:0",
                 "-c:v", "mpeg4", "-c:a", "aac", "-y", str(video)],
                check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                timeout=30)
            result = decoded_media_qa.inspect(str(video), sample_seconds=1)
            self.assertTrue(result["success"], result)
            # An audio track may intentionally end before the video (e.g. BGM).
            # The default is advisory; strict export contracts must opt in.
            self.assertTrue(result["passed"], result)
            warnings = [issue for issue in result["issues"]
                        if issue["code"] == "EXPORT_STREAM_ENDS_EARLY"
                        and issue["severity"] == "warning"]
            self.assertTrue(any(issue["stream_type"] == "audio"
                                and issue["missing_tail_seconds"] >= 2.9
                                for issue in warnings), result)
            strict = decoded_media_qa.inspect(
                str(video), sample_seconds=1, require_tracks_to_end=True)
            self.assertTrue(strict["success"], strict)
            self.assertFalse(strict["passed"], strict)
            self.assertTrue(any(issue["code"] == "EXPORT_STREAM_ENDS_EARLY"
                                and issue["stream_type"] == "audio"
                                and issue["severity"] == "error"
                                for issue in strict["issues"]), strict)


    def test_full_decode_rejects_empty_secondary_audio_stream(self):
        """One valid audio track must not conceal an empty second track."""
        with tempfile.TemporaryDirectory() as directory:
            video = Path(directory) / "two_audio_tracks.mkv"
            subprocess.run(
                ["ffmpeg", "-hide_banner", "-nostdin", "-loglevel", "error",
                 "-f", "lavfi", "-i", "color=c=blue:s=64x64:r=10:d=3",
                 "-f", "lavfi", "-i", "sine=frequency=440:duration=3",
                 "-itsoffset", "100", "-f", "lavfi", "-i",
                 "sine=frequency=880:duration=3",
                 "-map", "0:v", "-map", "1:a", "-map", "2:a",
                 "-t", "3", "-c:v", "mpeg4", "-c:a", "aac", "-y", str(video)],
                check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                timeout=30)
            # Some FFmpeg muxers may drop an entirely empty stream; only
            # exercise this regression when both declared tracks are retained.
            probe = subprocess.run(
                ["ffprobe", "-v", "error", "-show_entries",
                 "stream=codec_type", "-of", "json", str(video)],
                check=True, capture_output=True, text=True, timeout=15)
            import json
            streams = json.loads(probe.stdout)["streams"]
            if sum(s.get("codec_type") == "audio" for s in streams) != 2:
                self.skipTest("FFmpeg did not retain the empty second audio track")
            result = decoded_media_qa.inspect(str(video), sample_seconds=3,
                                              full_decode=True)
            self.assertTrue(result["success"], result)
            self.assertEqual(result["audio_streams"], 2)
            self.assertFalse(result["passed"], result)
            self.assertFalse(result["fully_decoded"], result)
            self.assertIn("EXPORT_FULL_DECODE_FAILED",
                          {issue["code"] for issue in result["issues"]})


if __name__ == "__main__":
    unittest.main()
