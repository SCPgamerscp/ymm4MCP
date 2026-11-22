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


    def test_strict_black_limit_includes_last_frames_display_time(self):
        # The final six frames are black. Their timestamps span 0.5s at
        # 10 FPS, but the last frame remains visible up to the 3.0s EOF.
        video = Path(self.directory.name) / "black-at-eof.mp4"
        subprocess.run([
            "ffmpeg", "-hide_banner", "-loglevel", "error", "-nostdin", "-y",
            "-f", "lavfi", "-i", "color=c=white:s=64x64:r=10:d=3",
            "-vf", "drawbox=c=black:t=fill:enable='gte(t,2.4)'",
            "-c:v", "mpeg4", "-q:v", "5", str(video),
        ], check=True, capture_output=True, timeout=30)
        options = dict(require_audio=False, sample_seconds=1)

        for scan in (False, True):
            with self.subTest(full_signal_scan=scan):
                report = qa.inspect(str(video), max_black_seconds=0.55,
                                    full_decode=scan, full_signal_scan=scan, **options)
                self.assertFalse(report["passed"], report)
                errors = [issue for issue in report["issues"]
                          if issue["code"] == "EXPORT_BLACK_DURATION_EXCEEDED"]
                self.assertEqual(len(errors), 1, report)
                self.assertAlmostEqual(errors[0]["video_start_seconds"], 2.4)
                self.assertAlmostEqual(errors[0]["video_end_seconds"], 3.0)
                self.assertAlmostEqual(errors[0]["duration_seconds"], 0.6)

        self.assertTrue(qa.inspect(str(video), max_black_seconds=0.6,
                                   full_decode=True, full_signal_scan=True, **options)["passed"])

    def test_full_scan_detects_trailing_blackout_longer_than_sample(self):
        video = Path(self.directory.name) / "long-black-at-eof.mp4"
        subprocess.run([
            "ffmpeg", "-hide_banner", "-loglevel", "error", "-nostdin", "-y",
            "-f", "lavfi", "-i", "color=c=white:s=64x64:r=10:d=6",
            "-vf", "drawbox=c=black:t=fill:enable='gte(t,4.1)'",
            "-c:v", "mpeg4", "-q:v", "5", str(video),
        ], check=True, capture_output=True, timeout=30)
        options = dict(require_audio=False, sample_seconds=1, max_black_seconds=1.85)
        sampled = qa.inspect(str(video), **options)
        self.assertTrue(sampled["passed"], sampled)

        scanned = qa.inspect(str(video), full_decode=True, full_signal_scan=True, **options)
        self.assertFalse(scanned["passed"], scanned)
        issue = next(issue for issue in scanned["issues"]
                     if issue["code"] == "EXPORT_BLACK_DURATION_EXCEEDED")
        self.assertAlmostEqual(issue["video_start_seconds"], 4.1)
        self.assertAlmostEqual(issue["video_end_seconds"], 6.0)
        self.assertAlmostEqual(issue["duration_seconds"], 1.9)

    def test_full_black_scan_keeps_tail_after_many_short_transitions(self):
        # More than 9,000 short-transition records previously filled the
        # stderr budget, even though those black frames are below the limit.
        video = Path(self.directory.name) / "short-transitions-and-black-tail.mp4"
        subprocess.run([
            "ffmpeg", "-hide_banner", "-loglevel", "error", "-nostdin", "-y",
            "-f", "lavfi", "-i", "color=c=white:s=16x16:r=120:d=126",
            "-vf", "drawbox=c=black:t=fill:enable='not(mod(n,3))+gte(t,124)'",
            "-c:v", "mpeg4", "-q:v", "5", str(video),
        ], check=True, capture_output=True, timeout=30)

        report = qa.inspect(str(video), require_audio=False, sample_seconds=1,
                            max_black_seconds=3, full_decode=True, full_signal_scan=True)
        self.assertTrue(report["passed"], report)
        self.assertEqual(report["signal_scope"], "full")
        self.assertEqual(len(report["black_frames"]), 1, report)
        event = report["black_frames"][0]
        self.assertAlmostEqual(event["video_start_seconds"], 124.0)
        self.assertAlmostEqual(event["video_end_seconds"], 126.0)

    def test_short_black_interval_that_ends_before_eof_does_not_extend(self):
        video = Path(self.directory.name) / "black-before-eof.mp4"
        subprocess.run([
            "ffmpeg", "-hide_banner", "-loglevel", "error", "-nostdin", "-y",
            "-f", "lavfi", "-i", "color=c=white:s=64x64:r=10:d=3",
            "-vf", "drawbox=c=black:t=fill:enable='gte(t,2.4)*lt(t,2.9)'",
            "-c:v", "mpeg4", "-q:v", "5", str(video),
        ], check=True, capture_output=True, timeout=30)
        report = qa.inspect(str(video), require_audio=False, sample_seconds=1,
                            max_black_seconds=0.55, full_decode=True, full_signal_scan=True)
        self.assertTrue(report["passed"], report)
        self.assertFalse(report["black_frames"], report)


if __name__ == "__main__":
    unittest.main()
