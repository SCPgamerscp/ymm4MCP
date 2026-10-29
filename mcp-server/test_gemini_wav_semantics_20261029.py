"""WAV format integrity and audio-evidence regression tests for Issue #61."""
import io
import json
import tempfile
import unittest
import wave
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock, patch

import gemini_video


class GeminiWavSemanticsTests(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.root = Path(tmp.name)
        (self.root / "0.png").write_bytes(b"synthetic-frame")
        self.audio = self.root / "audio.wav"
        self.client = MagicMock()
        self.client.models.generate_content.return_value = SimpleNamespace(
            text='{"summary":"ok","events":[]}')
        self.part = SimpleNamespace(
            from_bytes=lambda **kwargs: kwargs["data"],
            from_text=lambda **kwargs: kwargs["text"])

    def wav(self, *, channels=1, sample_width=2, rate=8000):
        stream = io.BytesIO()
        with wave.open(stream, "wb") as output:
            output.setnchannels(channels)
            output.setsampwidth(sample_width)
            output.setframerate(rate)
            output.writeframes(b"\x00" * (channels * sample_width * 2))
        return stream.getvalue()

    def analyze(self, payload=None, *, audio=True):
        if payload is not None:
            self.audio.write_bytes(payload)
        with patch.object(gemini_video, "is_available", return_value=(True, "ok")), \
             patch.object(gemini_video, "_client", return_value=self.client), \
             patch.object(gemini_video, "genai_types",
                          SimpleNamespace(Part=self.part)):
            return gemini_video.analyze_frames_dir(
                str(self.root), audio_path=str(self.audio) if audio else None)

    def reject(self, payload, fragment):
        self.client.models.generate_content.reset_mock()
        result = self.analyze(payload)
        self.assertFalse(result["success"], result)
        self.assertIn(fragment, result["error"])
        self.client.models.generate_content.assert_not_called()

    @staticmethod
    def field(payload, start, width, value):
        result = bytearray(payload)
        result[start:start + width] = value.to_bytes(width, "little")
        return bytes(result)

    def test_valid_pcm_and_float_wav_are_accepted(self):
        for channels, width in ((1, 1), (2, 2), (1, 3), (2, 4)):
            with self.subTest(channels=channels, width=width):
                self.assertTrue(self.analyze(self.wav(
                    channels=channels, sample_width=width))["success"])
        float_wav = self.field(self.wav(sample_width=4), 20, 2, 3)
        self.assertTrue(self.analyze(float_wav)["success"])

    def test_compressed_wav_codec_is_not_mislabelled_as_pcm(self):
        self.reject(self.field(self.wav(), 20, 2, 6), "unsupported WAV codec")

    def test_impossible_channel_counts_and_sample_rates_are_rejected(self):
        for start, width, value in ((22, 2, 0), (22, 2, 33),
                                    (24, 4, 0), (24, 4, 768001)):
            with self.subTest(field=start, value=value):
                self.reject(self.field(self.wav(), start, width, value),
                            "channel count or sample rate")

    def test_invalid_sample_precision_is_rejected(self):
        for bits in (0, 12, 64):
            with self.subTest(bits=bits):
                self.reject(self.field(self.wav(), 34, 2, bits),
                            "sample precision")
        self.reject(self.field(
            self.field(self.wav(), 20, 2, 3), 34, 2, 16),
            "sample precision")

    def test_byte_rate_and_block_alignment_must_match_pcm_layout(self):
        self.reject(self.field(self.wav(), 28, 4, 42), "block alignment")
        self.reject(self.field(self.wav(), 32, 2, 3), "block alignment")

    def test_partial_audio_sample_frame_is_rejected(self):
        payload = bytearray(self.wav())
        payload.extend(b"\x00\x00")  # extra sample byte and RIFF pad byte
        payload[40:44] = (5).to_bytes(4, "little")
        payload[4:8] = (len(payload) - 8).to_bytes(4, "little")
        self.reject(bytes(payload), "partial WAV sample frame")

    def test_data_chunk_cannot_precede_format_chunk(self):
        payload = self.wav()
        reordered = (payload[:12] + payload[36:] + payload[12:36])
        self.reject(reordered, "before fmt")

    def test_inconsistent_fmt_extension_is_rejected(self):
        payload = bytearray(self.wav())
        payload[16:20] = (18).to_bytes(4, "little")
        payload[36:36] = b"\x02\x00"  # cbSize claims two extra bytes
        payload[4:8] = (len(payload) - 8).to_bytes(4, "little")
        self.reject(bytes(payload), "extension length")
        payload[36:38] = b"\x00\x00"  # a legal empty WAVEFORMATEX extension
        self.assertTrue(self.analyze(bytes(payload))["success"])

    @staticmethod
    def extensible_wav(*, subtype=None, valid_bits=20):
        pcm_guid = bytes.fromhex("0100000000001000800000aa00389b71")
        fmt = (0xFFFE).to_bytes(2, "little")
        fmt += (2).to_bytes(2, "little")
        fmt += (48000).to_bytes(4, "little")
        fmt += (48000 * 6).to_bytes(4, "little")
        fmt += (6).to_bytes(2, "little")
        fmt += (24).to_bytes(2, "little")
        fmt += (22).to_bytes(2, "little")
        fmt += valid_bits.to_bytes(2, "little")
        fmt += (3).to_bytes(4, "little")
        fmt += subtype if subtype is not None else pcm_guid
        body = b"WAVE" + b"fmt " + len(fmt).to_bytes(4, "little") + fmt
        body += b"data" + (6).to_bytes(4, "little") + b"\x00" * 6
        return b"RIFF" + len(body).to_bytes(4, "little") + body

    def test_valid_extensible_pcm_is_accepted(self):
        self.assertTrue(self.analyze(self.extensible_wav())["success"])

    def test_extensible_channel_mask_matches_declared_channels(self):
        self.assertTrue(self.analyze(self.extensible_wav())["success"])
        payload = bytearray(self.extensible_wav())
        # The original mask has two speaker positions for two channels.
        payload[40:44] = (1).to_bytes(4, "little")
        self.reject(bytes(payload), "channel mask")
        payload[40:44] = (0).to_bytes(4, "little")
        self.assertTrue(self.analyze(bytes(payload))["success"])

    def test_excessive_wav_chunk_count_is_rejected(self):
        payload = self.wav()
        junk = b"JUNK" + (0).to_bytes(4, "little")
        body = payload[8:36] + junk * 4097 + payload[36:]
        oversized = b"RIFF" + len(body).to_bytes(4, "little") + body
        self.reject(oversized, "WAV chunk count limit")
        body = payload[8:36] + junk * 2 + payload[36:]
        normal = b"RIFF" + len(body).to_bytes(4, "little") + body
        self.assertTrue(self.analyze(normal)["success"])

    def test_unknown_extensible_subtype_and_invalid_bits_are_rejected(self):
        self.reject(self.extensible_wav(subtype=b"\x00" * 16),
                    "unsupported extensible WAV codec")
        self.reject(self.extensible_wav(valid_bits=25),
                    "invalid extensible WAV valid bits")

    def test_model_audio_claim_without_supplied_track_is_not_a_pass(self):
        for event in (
                {"time_sec": 0, "type": "sound", "audio": None},
                {"time_sec": 0, "type": "action", "audio": "loud crash"}):
            with self.subTest(event=event):
                self.client.models.generate_content.return_value = SimpleNamespace(
                    text=json.dumps({"summary": "ok", "events": [event]}))
                result = self.analyze(audio=False)
                self.assertFalse(result["success"], result)
                self.assertIn("without a supplied audio", result["error"])

    def test_visual_events_and_real_audio_are_not_blocked(self):
        event = {"time_sec": 0, "type": "action", "audio": None}
        self.client.models.generate_content.return_value = SimpleNamespace(
            text=json.dumps({"summary": "ok", "events": [event]}))
        self.assertTrue(self.analyze(audio=False)["success"])
        event["type"], event["audio"] = "sound", "voice"
        self.client.models.generate_content.return_value = SimpleNamespace(
            text=json.dumps({"summary": "ok", "events": [event]}))
        self.assertTrue(self.analyze(self.wav())["success"])


if __name__ == "__main__":
    unittest.main()
