"""Regression tests for capture ordering, file races and complete WAV QA (#61)."""
import io
import os
import tempfile
import unittest
import wave
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock, patch

import gemini_video


class CaptureRaceAndWavTests(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)
        self.png = self.root / "a.png"
        self.png.write_bytes(b"original-frame")
        available = patch.object(gemini_video, "is_available", return_value=(True, "ok"))
        available.start()
        self.addCleanup(available.stop)
        self.client = MagicMock()
        self.client.models.generate_content.return_value = SimpleNamespace(
            text='{"summary":"ok","events":[]}')
        self.uploaded_images = []

        def from_bytes(*, data, mime_type):
            if mime_type == "image/png":
                self.uploaded_images.append(data)
            return data

        self.part = SimpleNamespace(
            from_bytes=from_bytes,
            from_text=lambda **kwargs: kwargs["text"])

    def analyze(self, **kwargs):
        with patch.object(gemini_video, "_client", return_value=self.client), \
             patch.object(gemini_video, "genai_types",
                          SimpleNamespace(Part=self.part)):
            return gemini_video.analyze_frames_dir(str(self.root), **kwargs)

    def wav_bytes(self, samples=b"\x00\x00"):
        out = io.BytesIO()
        with wave.open(out, "wb") as wav:
            wav.setnchannels(1)
            wav.setsampwidth(2)
            wav.setframerate(8000)
            wav.writeframes(samples)
        return out.getvalue()

    def test_index_only_metadata_uses_capture_order_not_filename_order(self):
        (self.root / "z.png").write_bytes(b"first")
        self.png.write_bytes(b"second")
        result = self.analyze(frames_meta=[
            {"file": "a.png", "index": 1},
            {"file": "z.png", "index": 0},
        ])
        self.assertTrue(result["success"], result)
        self.assertEqual(self.uploaded_images, [b"first", b"second"])

    def test_frame_replaced_after_preflight_is_not_uploaded(self):
        def replace():
            self.png.write_bytes(b"modified-frame")
            st = self.png.stat()
            os.utime(self.png, ns=(st.st_atime_ns, st.st_mtime_ns + 1_000_000_000))
            return self.client

        with patch.object(gemini_video, "_client", side_effect=replace), \
             patch.object(gemini_video, "genai_types",
                          SimpleNamespace(Part=self.part)):
            result = gemini_video.analyze_frames_dir(str(self.root))
        self.assertFalse(result["success"], result)
        self.assertIn("changed", result["error"])
        self.client.models.generate_content.assert_not_called()

    def test_unchanged_frame_still_passes(self):
        result = self.analyze()
        self.assertTrue(result["success"], result)
        self.assertEqual(self.uploaded_images, [b"original-frame"])

    def test_header_only_wav_is_not_audio_qa(self):
        audio = self.root / "audio.wav"
        audio.write_bytes(b"RIFF" + (4).to_bytes(4, "little") + b"WAVE")
        result = self.analyze(audio_path=str(audio))
        self.assertFalse(result["success"], result)
        self.assertIn("lacks a WAV fmt", result["error"])
        self.client.models.generate_content.assert_not_called()

    def test_empty_audio_data_chunk_is_not_a_complete_capture(self):
        audio = self.root / "audio.wav"
        audio.write_bytes(self.wav_bytes(b""))
        result = self.analyze(audio_path=str(audio))
        self.assertFalse(result["success"], result)
        self.assertIn("empty WAV data", result["error"])
        self.client.models.generate_content.assert_not_called()

    def test_truncated_wav_chunk_is_rejected(self):
        audio = self.root / "audio.wav"
        payload = self.wav_bytes()
        # Make the fmt chunk claim more bytes than the RIFF payload contains.
        changed = bytearray(payload)
        changed[16:20] = (len(payload) + 100).to_bytes(4, "little")
        audio.write_bytes(changed)
        result = self.analyze(audio_path=str(audio))
        self.assertFalse(result["success"], result)
        self.assertIn("truncated WAV chunk", result["error"])
        self.client.models.generate_content.assert_not_called()

    def test_duplicate_data_chunks_are_rejected(self):
        audio = self.root / "audio.wav"
        payload = self.wav_bytes()
        duplicate = b"data" + (2).to_bytes(4, "little") + b"\x00\x00"
        changed = bytearray(payload + duplicate)
        changed[4:8] = (len(changed) - 8).to_bytes(4, "little")
        audio.write_bytes(changed)
        result = self.analyze(audio_path=str(audio))
        self.assertFalse(result["success"], result)
        self.assertIn("duplicate", result["error"])
        self.client.models.generate_content.assert_not_called()

    def test_valid_complete_wav_still_passes(self):
        audio = self.root / "audio.wav"
        audio.write_bytes(self.wav_bytes())
        result = self.analyze(audio_path=str(audio))
        self.assertTrue(result["success"], result)
        self.client.models.generate_content.assert_called_once()

    def test_audio_changed_after_preflight_is_rejected(self):
        audio = self.root / "audio.wav"
        audio.write_bytes(self.wav_bytes())
        real_open = open

        def change_before_open(path, mode="r", *args, **kwargs):
            if str(path) == str(audio) and mode == "rb":
                audio.write_bytes(self.wav_bytes(b"\x01\x02"))
                st = audio.stat()
                os.utime(audio, ns=(st.st_atime_ns, st.st_mtime_ns + 1_000_000_000))
            return real_open(path, mode, *args, **kwargs)

        with patch("builtins.open", side_effect=change_before_open):
            result = self.analyze(audio_path=str(audio))
        self.assertFalse(result["success"], result)
        self.assertIn("changed", result["error"])
        self.client.models.generate_content.assert_not_called()


if __name__ == "__main__":
    unittest.main()
