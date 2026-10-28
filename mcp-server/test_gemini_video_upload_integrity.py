"""Regression tests for bounded, stable Gemini File API video analysis."""
import tempfile
import unittest

import httpx
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock, patch

import gemini_video


class FileApiError(Exception):
    """Match google.genai.errors.APIError: HTTP status is exposed as code."""
    def __init__(self, code):
        super().__init__(f"HTTP {code}")
        self.code = code


class GeminiVideoUploadIntegrityTests(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.video = Path(temp.name) / "clip.mp4"
        self.video.write_bytes(b"original video")
        self.client = MagicMock()
        self.client.files.upload.return_value = SimpleNamespace(name="files/uploaded")
        self.client.models.generate_content.return_value = SimpleNamespace(
            text='{"summary":"ok","events":[]}'
        )

    def analyze(self):
        with patch.object(gemini_video, "is_available", return_value=(True, "ok")), \
             patch.object(gemini_video, "_client", return_value=self.client), \
             patch.object(gemini_video, "_wait_file_active",
                          return_value=SimpleNamespace(name="files/uploaded")):
            return gemini_video.analyze_video_file(str(self.video))

    def test_upload_budget_and_empty_video_fail_before_client_creation(self):
        with patch.object(gemini_video, "MAX_VIDEO_UPLOAD_BYTES", 5):
            result = self.analyze()
        self.assertFalse(result["success"])
        self.assertIn("size limit", result["error"])
        self.client.files.upload.assert_not_called()
        self.video.write_bytes(b"")
        result = self.analyze()
        self.assertFalse(result["success"])
        self.client.files.upload.assert_not_called()

    def test_video_replaced_during_upload_never_reaches_model(self):
        def mutate(file):
            self.video.write_bytes(b"changed content")
            return SimpleNamespace(name="files/uploaded")
        self.client.files.upload.side_effect = mutate
        result = self.analyze()
        self.assertFalse(result["success"])
        self.assertIn("changed during Gemini upload", result["error"])
        self.client.models.generate_content.assert_not_called()
        self.client.files.delete.assert_called_once_with(name="files/uploaded")

    def test_video_replaced_during_generation_is_not_certified(self):
        def mutate(**kwargs):
            self.video.write_bytes(b"new content")
            return SimpleNamespace(text='{"summary":"ok","events":[]}')
        self.client.models.generate_content.side_effect = mutate
        result = self.analyze()
        self.assertFalse(result["success"])
        self.assertIn("changed during Gemini analysis", result["error"])
        self.client.files.delete.assert_called_once_with(name="files/uploaded")

    def test_unchanged_video_succeeds_and_upload_is_deleted(self):
        result = self.analyze()
        self.assertTrue(result["success"])
        self.client.files.delete.assert_called_once_with(name="files/uploaded")


class GeminiFilePollingRetryTests(unittest.TestCase):
    def setUp(self):
        self.client = MagicMock()
        self.uploaded = SimpleNamespace(name="files/uploaded")

    def test_429_and_503_retry_then_active(self):
        self.client.files.get.side_effect = [
            FileApiError(429), FileApiError(503),
            SimpleNamespace(state=SimpleNamespace(name="ACTIVE")),
        ]
        with patch.object(gemini_video.time, "monotonic", return_value=0), \
             patch.object(gemini_video.time, "sleep") as sleep:
            result = gemini_video._wait_file_active(self.client, self.uploaded)
        self.assertEqual(result.state.name, "ACTIVE")
        self.assertEqual([call.args[0] for call in sleep.call_args_list], [3, 6])
        self.assertEqual(self.client.files.get.call_count, 3)

    def test_each_poll_has_remaining_request_timeout(self):
        self.client.files.get.side_effect = [
            SimpleNamespace(state=SimpleNamespace(name="PROCESSING")),
            SimpleNamespace(state=SimpleNamespace(name="ACTIVE")),
        ]
        with patch.object(gemini_video.time, "monotonic",
                          side_effect=[0, 0, 1, 5, 5, 6]), \
             patch.object(gemini_video.time, "sleep") as sleep:
            result = gemini_video._wait_file_active(
                self.client, self.uploaded, timeout_sec=10)
        self.assertEqual(result.state.name, "ACTIVE")
        self.assertEqual(sleep.call_args.args[0], 3)
        requests = self.client.files.get.call_args_list
        self.assertEqual(len(requests), 2)
        self.assertEqual(
            [call.kwargs["config"]["http_options"]["timeout"] for call in requests],
            [10000, 5000],
        )

    def test_submillisecond_budget_skips_network_request(self):
        with patch.object(gemini_video.time, "monotonic",
                          side_effect=[0, 9.9995]):
            with self.assertRaises(TimeoutError):
                gemini_video._wait_file_active(
                    self.client, self.uploaded, timeout_sec=10)
        self.client.files.get.assert_not_called()

    def test_unauthorized_is_not_retried(self):
        self.client.files.get.side_effect = FileApiError(401)
        with patch.object(gemini_video.time, "sleep") as sleep:
            with self.assertRaises(FileApiError):
                gemini_video._wait_file_active(self.client, self.uploaded)
        self.assertEqual(self.client.files.get.call_count, 1)
        sleep.assert_not_called()

    def test_retry_respects_wall_clock_deadline(self):
        self.client.files.get.side_effect = FileApiError(503)
        with patch.object(gemini_video.time, "monotonic",
                          side_effect=[0, 0, 9.5, 10.1]), \
             patch.object(gemini_video.time, "sleep") as sleep:
            with self.assertRaises(TimeoutError):
                gemini_video._wait_file_active(self.client, self.uploaded, timeout_sec=10)
        sleep.assert_called_once_with(0.5)
        self.assertEqual(self.client.files.get.call_count, 1)

    def test_active_response_after_deadline_is_rejected(self):
        self.client.files.get.return_value = SimpleNamespace(
            state=SimpleNamespace(name="ACTIVE"))
        with patch.object(gemini_video.time, "monotonic",
                          side_effect=[0, 0, 10.1]), \
             patch.object(gemini_video.time, "sleep") as sleep:
            with self.assertRaisesRegex(TimeoutError, "polling timed out"):
                gemini_video._wait_file_active(self.client, self.uploaded, timeout_sec=10)
        self.assertEqual(self.client.files.get.call_count, 1)
        sleep.assert_not_called()

    def test_transport_error_retried_but_bad_request_not_retried(self):
        self.client.files.get.side_effect = [
            httpx.ConnectError("temporary network failure"),
            SimpleNamespace(state=SimpleNamespace(name="ACTIVE")),
        ]
        with patch.object(gemini_video.time, "monotonic", return_value=0), \
             patch.object(gemini_video.time, "sleep") as sleep:
            self.assertEqual(
                gemini_video._wait_file_active(self.client, self.uploaded).state.name,
                "ACTIVE")
        sleep.assert_called_once_with(3)
        self.assertEqual(self.client.files.get.call_count, 2)

        self.client.files.get.reset_mock()
        self.client.files.get.side_effect = FileApiError(400)
        with patch.object(gemini_video.time, "sleep") as sleep:
            with self.assertRaises(FileApiError):
                gemini_video._wait_file_active(self.client, self.uploaded)
        self.assertEqual(self.client.files.get.call_count, 1)
        sleep.assert_not_called()


if __name__ == "__main__":
    unittest.main()
