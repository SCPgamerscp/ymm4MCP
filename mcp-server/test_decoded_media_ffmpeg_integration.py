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
            result = decoded_media_qa.inspect(str(video), sample_seconds=4)
            self.assertTrue(result["success"], result)
            self.assertTrue(result["passed"], result)  # warnings are not errors
            self.assertEqual(result["signal_scope"], "sampled_only")
            self.assertGreaterEqual(len(result["black_frames"]), 1)
            self.assertGreaterEqual(len(result["long_silences"]), 1)
            self.assertTrue(all(issue["severity"] == "warning"
                                for issue in result["issues"]))


if __name__ == "__main__":
    unittest.main()
