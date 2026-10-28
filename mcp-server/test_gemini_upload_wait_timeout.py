"""Gemini File API polling uses elapsed wall time, not sleep counts."""
from types import SimpleNamespace
import unittest
from unittest.mock import MagicMock, patch

import gemini_video


def file_state(name):
    return SimpleNamespace(state=SimpleNamespace(name=name))


class GeminiUploadWaitTimeoutTests(unittest.TestCase):
    def setUp(self):
        self.client = MagicMock()
        self.uploaded = SimpleNamespace(name="files/uploaded")

    def test_slow_get_calls_count_toward_timeout(self):
        self.client.files.get.side_effect = [
            file_state("PROCESSING"), file_state("PROCESSING"),
        ]
        with patch.object(gemini_video.time, "monotonic", side_effect=[0, 0, 4, 7, 11]), \
             patch.object(gemini_video.time, "sleep") as sleep:
            with self.assertRaises(TimeoutError):
                gemini_video._wait_file_active(self.client, self.uploaded, timeout_sec=10)
        self.assertEqual(self.client.files.get.call_count, 2)
        sleep.assert_called_once_with(3)

    def test_last_sleep_is_limited_to_remaining_time(self):
        self.client.files.get.side_effect = [
            file_state("PROCESSING"), file_state("PROCESSING"),
        ]
        with patch.object(gemini_video.time, "monotonic", side_effect=[0, 0, 9.5, 10.1]), \
             patch.object(gemini_video.time, "sleep") as sleep:
            with self.assertRaises(TimeoutError):
                gemini_video._wait_file_active(self.client, self.uploaded, timeout_sec=10)
        self.assertAlmostEqual(sleep.call_args.args[0], 0.5)
        self.assertEqual(self.client.files.get.call_count, 1)

    def test_active_and_failed_states_keep_existing_behavior(self):
        self.client.files.get.return_value = file_state("ACTIVE")
        with patch.object(gemini_video.time, "monotonic", return_value=0), \
             patch.object(gemini_video.time, "sleep") as sleep:
            self.assertEqual(gemini_video._wait_file_active(
                self.client, self.uploaded).state.name, "ACTIVE")
            sleep.assert_not_called()
        self.client.files.get.return_value = file_state("FAILED")
        with patch.object(gemini_video.time, "monotonic", return_value=0):
            with self.assertRaises(RuntimeError):
                gemini_video._wait_file_active(self.client, self.uploaded)


if __name__ == "__main__":
    unittest.main()
