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


    def test_short_audio_stream_fails_tail_sample_even_when_video_is_longer(self):
        """FFmpeg without -abort_on empty_output incorrectly returns exit 0."""
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
            self.assertFalse(result["passed"], result)
            failures = [issue for issue in result["issues"]
                        if issue["code"] == "EXPORT_DECODE_FAILED"]
            self.assertTrue(any(issue["stream_type"] == "audio"
                                and issue["start_seconds"] >= 3.0
                                for issue in failures), result)


if __name__ == "__main__":
    unittest.main()
