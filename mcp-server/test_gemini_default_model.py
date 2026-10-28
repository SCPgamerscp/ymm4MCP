"""Regression tests for the bundled Gemini video analysis model selection."""
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import gemini_video


class BundledGeminiModelTests(unittest.TestCase):
    def configured(self, raw: bytes | None) -> str:
        with tempfile.TemporaryDirectory() as folder:
            if raw is not None:
                (Path(folder) / "config.json").write_bytes(raw)
            with patch.object(gemini_video, "__file__", str(Path(folder) / "gemini_video.py")):
                return gemini_video._get_default_model()

    def test_bundled_default_uses_supported_stable_video_model(self):
        config = json.loads((Path(gemini_video.__file__).parent / "config.json").read_text(
            encoding="utf-8"))
        self.assertEqual(config["default_model"], "gemini-3.8-flash")
        self.assertEqual(gemini_video.DEFAULT_MODEL, config["default_model"])

    def test_missing_or_invalid_config_uses_supported_fallback(self):
        fallback = gemini_video.DEFAULT_GEMINI_MODEL
        self.assertEqual(self.configured(None), fallback)
        for invalid in (b"{}", b"null", b"[]", b"{", b"\xff",
                        b'{"default_model": null}', b'{"default_model": 123}',
                        b'{"default_model": ""}', b'{"default_model": "   "}',
                        b'{"default_model": "gemini-3.8-flash\\n"}',
                        b'{"default_model": "../models/gemini"}',
                        b'{"default_model": "' + b"x" * 129 + b'"}',
                        b" " * (gemini_video.MAX_MODEL_CONFIG_BYTES + 1)):
            with self.subTest(config=invalid[:60]):
                self.assertEqual(self.configured(invalid), fallback)

    def test_explicit_valid_model_override_remains_supported(self):
        for name in ("gemini-3.6-flash", "models/gemini-3.8-flash",
                     "custom-model_v2.0"):
            with self.subTest(name=name):
                self.assertEqual(self.configured(json.dumps(
                    {"default_model": name}).encode("utf-8")), name)

    def test_duplicate_model_key_falls_back_instead_of_using_last_value(self):
        raw = (b'{"default_model":"gemini-3.8-flash",'
               b'"default_model":"custom-model_v2.0"}')
        self.assertEqual(self.configured(raw), gemini_video.DEFAULT_GEMINI_MODEL)

    def test_duplicate_nested_metadata_key_also_falls_back(self):
        raw = (b'{"default_model":"custom-model_v2.0",'
               b'"metadata":{"name":"a","name":"b"}}')
        self.assertEqual(self.configured(raw), gemini_video.DEFAULT_GEMINI_MODEL)

    def test_ambiguous_model_path_segments_are_rejected(self):
        for model in ("models//gemini-3.8-flash", "models/../gemini-3.8-flash",
                      "models/./gemini-3.8-flash", "gemini-3.8-flash/"):
            with self.subTest(model=model):
                self.assertEqual(self.configured(json.dumps(
                    {"default_model": model}).encode("utf-8")),
                    gemini_video.DEFAULT_GEMINI_MODEL)

    def test_deeply_nested_config_cannot_crash_model_initialization(self):
        # json.loads can raise RecursionError even for a <64 KiB file.
        for raw in (b"[" * 60000, b'{"metadata":' + b"[" * 2000 + b"0" + b"]" * 2000 + b"}"):
            with self.subTest(length=len(raw)):
                self.assertEqual(self.configured(raw), gemini_video.DEFAULT_GEMINI_MODEL)

    def test_utf8_bom_config_is_accepted(self):
        self.assertEqual(self.configured(b"\xef\xbb\xbf" +
                         b'{"default_model": "gemini-3.8-flash"}'),
                         "gemini-3.8-flash")


if __name__ == "__main__":
    unittest.main()
