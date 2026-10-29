"""Regression tests for Gemini prompt and filename safety in automated QA (#61)."""
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock, patch

import gemini_video


class GeminiPromptSafetyTests(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)
        (self.root / "frame_00000.png").write_bytes(b"synthetic-frame")
        self.availability = patch.object(
            gemini_video, "is_available", return_value=(True, "ok"))
        self.availability.start()
        self.addCleanup(self.availability.stop)

    def test_windows_png_names_reject_bidi_controls_and_surrogates(self):
        for filename in ("frame\u202Egnp.png", "frame\u2066.png",
                         "frame\u200F.png", "frame\u061C.png",
                         "frame\ud800.png"):
            with self.subTest(filename=ascii(filename)):
                self.assertFalse(gemini_video._valid_capture_png_name(filename))
                with patch.object(gemini_video, "_client",
                                  side_effect=AssertionError("no network")):
                    result = gemini_video.analyze_frames_dir(
                        str(self.root), frames_meta=[{"file": filename}])
                self.assertFalse(result["success"], result)
                self.assertIn("invalid PNG filename", result["error"])
        self.assertTrue(gemini_video._valid_capture_png_name("字幕_01.png"))

    def test_video_file_rejects_invalid_instruction_before_upload(self):
        video = self.root / "sample.mp4"
        video.write_bytes(b"test-video")
        for instruction in (42, [], "x" * 32769, "bad\x00prompt",
                            "bad\udfffprompt"):
            with self.subTest(instruction=ascii(instruction)[:50]):
                with patch.object(gemini_video, "_client",
                                  side_effect=AssertionError("no network")):
                    result = gemini_video.analyze_video_file(
                        str(video), event_instruction=instruction)
                self.assertFalse(result["success"], result)
                self.assertIn("event_instruction", result["error"])
                self.assertEqual(result["source"], "file")

    def test_instruction_validation_preserves_normal_text(self):
        self.assertTrue(gemini_video._valid_event_instruction(""))
        self.assertTrue(gemini_video._valid_event_instruction("字幕と効果音"))
        self.assertTrue(gemini_video._valid_event_instruction("a" * 32768))
        self.assertFalse(gemini_video._valid_event_instruction("a" * 32769))

    def _run_frames(self, audio_path=None):
        client = MagicMock()
        client.models.generate_content.return_value = SimpleNamespace(
            text='{"summary":"ok","events":[]}')
        parts = SimpleNamespace(
            from_bytes=lambda **kwargs: kwargs["data"],
            from_text=lambda **kwargs: kwargs["text"])
        with patch.object(gemini_video, "_client", return_value=client), \
             patch.object(gemini_video, "genai_types",
                          SimpleNamespace(Part=parts)):
            result = gemini_video.analyze_frames_dir(
                str(self.root), audio_path=audio_path)
        self.assertTrue(result["success"], result)
        return client.models.generate_content.call_args.kwargs["contents"][-1]

    def test_video_only_prompt_does_not_claim_audio_was_supplied(self):
        prompt = self._run_frames()
        self.assertIn("音声データは提供されていません", prompt)
        self.assertIn("音響イベントを観測したと推測して報告しない", prompt)
        self.assertNotIn("その区間の音声です", prompt)

    def test_audio_present_prompt_does_not_warn_audio_missing(self):
        audio = self.root / "audio.wav"
        # A minimal RIFF/WAVE envelope sufficient for this prompt-path test.
        audio.write_bytes(b"RIFF" + (4).to_bytes(4, "little") + b"WAVE")
        prompt = self._run_frames(str(audio))
        self.assertIn("その区間の音声です", prompt)
        self.assertNotIn("音声データは提供されていません", prompt)


if __name__ == "__main__":
    unittest.main()
