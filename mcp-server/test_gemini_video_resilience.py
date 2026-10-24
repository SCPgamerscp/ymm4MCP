"""Gemini analysis failures must not masquerade as successful QA or leak uploads."""
import json
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock, patch

import gemini_video


def response(**overrides):
    payload = {"summary": "検出結果", "events": []}
    payload.update(overrides)
    return json.dumps(payload, ensure_ascii=False)


class GeminiAnalysisSafetyTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.video = Path(self.temp.name) / "clip.mp4"
        self.video.write_bytes(b"fake video")
        self.client = MagicMock()
        self.client.files.upload.return_value = SimpleNamespace(name="files/uploaded")
        self.client.models.generate_content.return_value = SimpleNamespace(text=response())

    def analyze_file(self):
        with patch.object(gemini_video, "is_available", return_value=(True, "ok")), \
             patch.object(gemini_video, "_client", return_value=self.client), \
             patch.object(gemini_video, "_wait_file_active",
                          return_value=SimpleNamespace(name="files/uploaded")):
            return gemini_video.analyze_video_file(str(self.video))

    def test_success_deletes_upload_and_keeps_trusted_metadata(self):
        self.client.models.generate_content.return_value = SimpleNamespace(text=response(
            success=False, source="untrusted", video_path="untrusted", model="untrusted"))
        result = self.analyze_file()
        self.assertTrue(result["success"])
        self.assertEqual(result["source"], "file")
        self.assertEqual(result["video_path"], str(self.video))
        self.assertEqual(result["model"], gemini_video.DEFAULT_MODEL)
        self.client.files.delete.assert_called_once_with(name="files/uploaded")

    def test_generate_failure_also_deletes_upload(self):
        self.client.models.generate_content.side_effect = RuntimeError("generation failed")
        result = self.analyze_file()
        self.assertFalse(result["success"])
        self.assertIn("generation failed", result["error"])
        self.client.files.delete.assert_called_once_with(name="files/uploaded")

    def test_wait_failure_also_deletes_upload(self):
        with patch.object(gemini_video, "is_available", return_value=(True, "ok")), \
             patch.object(gemini_video, "_client", return_value=self.client), \
             patch.object(gemini_video, "_wait_file_active",
                          side_effect=TimeoutError("processing timeout")):
            result = gemini_video.analyze_video_file(str(self.video))
        self.assertFalse(result["success"])
        self.client.files.delete.assert_called_once_with(name="files/uploaded")

    def test_invalid_model_output_is_failure_and_upload_is_deleted(self):
        self.client.models.generate_content.return_value = SimpleNamespace(text="not JSON")
        result = self.analyze_file()
        self.assertFalse(result["success"])
        self.assertIn("JSON解析失敗", result["error"])
        self.client.files.delete.assert_called_once_with(name="files/uploaded")

    def test_client_creation_failure_is_reported(self):
        with patch.object(gemini_video, "is_available", return_value=(True, "ok")), \
             patch.object(gemini_video, "_client", side_effect=RuntimeError("bad credentials")):
            result = gemini_video.analyze_video_file(str(self.video))
        self.assertFalse(result["success"])
        self.assertIn("bad credentials", result["error"])

    def test_rejects_non_object_and_malformed_event_lists(self):
        for output in ('[]', '{"summary":"ok","events":{}}',
                       '{"summary":"ok","events":[123]}',
                       '{"summary":123,"events":[]}'):
            with self.subTest(output=output):
                self.assertIn("_parse_error", gemini_video._extract_json(output))

    def test_frames_failure_does_not_report_success(self):
        (Path(self.temp.name) / "0001.png").write_bytes(b"fake png")
        self.client.models.generate_content.return_value = SimpleNamespace(text="not JSON")
        part = SimpleNamespace(
            from_bytes=lambda **kwargs: SimpleNamespace(inline_data=kwargs["data"]),
            from_text=lambda **kwargs: SimpleNamespace(inline_data=None))
        with patch.object(gemini_video, "is_available", return_value=(True, "ok")), \
             patch.object(gemini_video, "_client", return_value=self.client), \
             patch.object(gemini_video, "genai_types", SimpleNamespace(Part=part)):
            result = gemini_video.analyze_frames_dir(self.temp.name)
        self.assertFalse(result["success"])
        self.assertEqual(result["source"], "frames_dir")
        self.assertIn("JSON解析失敗", result["error"])

    def test_frames_success_keeps_trusted_metadata(self):
        (Path(self.temp.name) / "0001.png").write_bytes(b"fake png")
        self.client.models.generate_content.return_value = SimpleNamespace(
            text=response(success=False, source="spoofed", frames_dir="spoofed", model="spoofed"))
        part = SimpleNamespace(
            from_bytes=lambda **kwargs: SimpleNamespace(inline_data=kwargs["data"]),
            from_text=lambda **kwargs: SimpleNamespace(inline_data=None))
        with patch.object(gemini_video, "is_available", return_value=(True, "ok")), \
             patch.object(gemini_video, "_client", return_value=self.client), \
             patch.object(gemini_video, "genai_types", SimpleNamespace(Part=part)):
            result = gemini_video.analyze_frames_dir(self.temp.name)
        self.assertTrue(result["success"])
        self.assertEqual(result["source"], "frames_dir")
        self.assertEqual(result["frames_dir"], self.temp.name)
        self.assertEqual(result["analyzed_images"], 1)
        self.assertEqual(result["model"], gemini_video.DEFAULT_MODEL)


if __name__ == "__main__":
    unittest.main()
