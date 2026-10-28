"""End-to-end deadline regression coverage for authenticated status checks (#22)."""
import itertools
import unittest
from unittest.mock import MagicMock, patch

import check_connection


class TotalDeadlineTests(unittest.TestCase):
    PAYLOAD = b'{"status":"running","port":8765,"version":"4.56.1.1"}'

    def probe(self, clock, chunks=None):
        client = MagicMock()
        client.__enter__.return_value = client
        rejected = MagicMock()
        rejected.__enter__.return_value = MagicMock(status_code=401)
        accepted = MagicMock()
        response = MagicMock(status_code=200)
        response.headers = {}
        response.iter_bytes.return_value = [self.PAYLOAD] if chunks is None else chunks
        accepted.__enter__.return_value = response
        client.stream.side_effect = [rejected, accepted]
        times = itertools.chain(clock, itertools.repeat(clock[-1]))
        with patch.object(check_connection, "connection_settings", return_value=(
                "http://127.0.0.1:8765/api", {"X-Ymm4-Token": "secret"})), \
             patch.object(check_connection.httpx, "Client", return_value=client), \
             patch.object(check_connection.time, "monotonic", side_effect=times):
            try:
                return check_connection.check()
            finally:
                self.last_client = client
                self.last_response = response

    def test_expired_negative_probe_never_sends_real_token(self):
        with self.assertRaisesRegex(ValueError, "タイムアウト"):
            self.probe([0.0, 11.0])
        self.assertEqual(self.last_client.stream.call_count, 1)

    def test_expired_positive_headers_are_rejected_before_body_read(self):
        with self.assertRaisesRegex(ValueError, "タイムアウト"):
            self.probe([0.0, 1.0, 11.0])
        self.last_response.iter_bytes.assert_not_called()

    def test_slow_eof_without_data_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "タイムアウト"):
            self.probe([0.0, 1.0, 2.0, 11.0], chunks=[])
        self.last_response.iter_bytes.assert_called_once_with()

    def test_slow_eof_after_valid_body_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "タイムアウト"):
            self.probe([0.0, 1.0, 2.0, 3.0, 11.0])

    def test_slow_stream_close_after_valid_body_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "タイムアウト"):
            self.probe([0.0, 1.0, 2.0, 3.0, 4.0, 11.0])

    def test_timely_probes_and_stream_close_succeed(self):
        result = self.probe([0.0, 1.0, 2.0, 3.0, 4.0, 5.0])
        self.assertEqual(result["port"], 8765)


if __name__ == "__main__":
    unittest.main()
