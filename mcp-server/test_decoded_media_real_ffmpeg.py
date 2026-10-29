"""Real FFmpeg/ffprobe end-to-end regression for completed-export QA (#61).

Mocked subprocess tests cannot detect broken filter options, incompatible
stream selectors, or incorrect parsing of actual FFmpeg diagnostic output.
This creates a tiny disposable MP4; it never needs a running YMM4 instance.
"""
import shutil
import subprocess
import tempfile
import unittest
from pathlib import Path

import decoded_media_qa as qa


@unittest.skipUnless(shutil.which("ffmpeg") and shutil.which("ffprobe"),
                     "real media QA requires ffmpeg and ffprobe on PATH")
class RealFfmpegExportQaTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.directory = tempfile.TemporaryDirectory()
        cls.addClassCleanup(cls.directory.cleanup)
        cls.video = Path(cls.directory.name) / "black-and-silent.mp4"
        # MPEG-4 Part 2 and AAC are widely available in stock FFmpeg builds.
        # All-black video and silent stereo audio deliberately exercise both
        # signal detectors without relying on platform-specific encoders.
        command = [
            "ffmpeg", "-hide_banner", "-loglevel", "error", "-nostdin", "-y",
            "-f", "lavfi", "-i", "color=c=black:s=64x64:r=10:d=3",
            "-f", "lavfi", "-i", "anullsrc=r=48000:cl=stereo",
            "-t", "3", "-c:v", "mpeg4", "-q:v", "5", "-c:a", "aac",
            "-movflags", "+faststart", str(cls.video),
        ]
        subprocess.run(command, check=True, capture_output=True, timeout=30)
        if cls.video.stat().st_size <= 0:
            raise AssertionError("FFmpeg did not create a media fixture")

    def inspect(self, **kwargs):
        options = dict(
            sample_seconds=3, expected_duration_seconds=3.0,
            duration_tolerance_seconds=0.25,
            expected_width=64, expected_height=64, expected_fps=10,
            fps_tolerance=0, expected_audio_sample_rate=48000,
            expected_audio_channels=2, expected_video_streams=1,
            expected_audio_streams=1,
        )
        options.update(kwargs)
        return qa.inspect(str(self.video), **options)

    def test_real_full_decode_and_black_silence_detection(self):
        report = self.inspect(full_decode=True)
        self.assertTrue(report["passed"], report)
        self.assertTrue(report["fully_decoded"], report)
        self.assertEqual(report["verification_scope"], "full_audio_video_decode")
        self.assertEqual(report["resolution"], "64x64")
        self.assertEqual(report["video_streams"], 1)
        self.assertEqual(report["audio_streams"], 1)
        # Real filter logs (not fabricated mock strings) must be understood.
        self.assertTrue(report["black_frames"], report)
        self.assertTrue(report["long_silences"], report)
        self.assertIn("EXPORT_SAMPLE_BLACK_FRAMES",
                      {issue["code"] for issue in report["issues"]})
        self.assertIn("EXPORT_SAMPLE_LONG_SILENCE",
                      {issue["code"] for issue in report["issues"]})

    def test_real_media_expected_format_mismatch_blocks_acceptance(self):
        report = self.inspect(expected_width=128)
        self.assertFalse(report["passed"], report)
        self.assertIn("EXPORT_VIDEO_FORMAT_MISMATCH",
                      {issue["code"] for issue in report["issues"]})
        self.assertFalse(report["fully_decoded"])


    def test_full_signal_scan_detects_blackout_between_sample_windows(self):
        # With 2s samples at 0, 7 and 14, a black segment at 4-6s
        # is invisible to sampling even when full_decode=True.
        # The opt-in signal pass must find it in the *entire* decoded movie.
        video = Path(self.directory.name) / "middle-blackout.mp4"
        subprocess.run([
            "ffmpeg", "-hide_banner", "-loglevel", "error", "-nostdin", "-y",
            "-f", "lavfi", "-i", "color=c=white:s=64x64:r=10:d=16",
            "-vf", "drawbox=c=black:t=fill:enable='between(t,4,6)'",
            "-c:v", "mpeg4", "-q:v", "5", str(video),
        ], check=True, capture_output=True, timeout=30)
        options = dict(require_audio=False, sample_seconds=2, full_decode=True)
        sampled = qa.inspect(str(video), **options)
        self.assertTrue(sampled["passed"], sampled)
        self.assertEqual(sampled["signal_scope"], "sampled_only")
        self.assertFalse(any(3.5 <= event["video_start_seconds"] <= 6.5
                             for event in sampled["black_frames"]), sampled)

        scanned = qa.inspect(str(video), full_signal_scan=True, **options)
        self.assertTrue(scanned["passed"], scanned)
        self.assertTrue(scanned["fully_decoded"], scanned)
        self.assertEqual(scanned["signal_scope"], "full")
        self.assertEqual(scanned["samples"], sampled["samples"])
        self.assertTrue(any(3.5 <= event["video_start_seconds"] <= 4.5
                            and event["video_end_seconds"] >= 6
                            for event in scanned["black_frames"]), scanned)


if __name__ == "__main__":
    unittest.main()
