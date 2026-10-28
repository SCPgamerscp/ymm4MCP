"""Regression tests for status media-type integrity and total stream deadline (#22)."""
import itertools
import unittest
from unittest.mock import MagicMock, patch

import check_connection


class StatusMediaTypeAndDeadlineTests(unittest.TestCase):
    PAYLOAD = b'{"status":"running","port":8765,"version":"4.56.1.1"}'

    def probe(self, headers=None, chunks=None, clock=None):
        response = MagicMock(status_code=200)
        response.headers = {} if headers is None else headers
        response.iter_bytes.return_value = [self.PAYLOAD] if chunks is None else chunks
        client = MagicMock()
        client.__enter__.return_value = client
        rejected = MagicMock()
        rejected.__enter__.return_value = MagicMock(status_code=401)
        accepted = MagicMock()
        accepted.__enter__.return_value = response
        client.stream.side_effect = [rejected, accepted]
        with patch.object(check_connection, "connection_settings", return_value=(
                "http://127.0.0.1:8765/api", {"X-Ymm4-Token": "secret"})), \
             patch.object(check_connection.httpx, "Client", return_value=client):
            if clock is None:
                result = check_connection.check()
            else:
                with patch.object(check_connection.time, "monotonic", side_effect=itertools.chain(clock, itertools.repeat(clock[-1]))):
                    result = check_connection.check()
        return result, response

    def test_rejects_declared_non_json_response_without_reading_body(self):
        for media_type in ("text/html", "text/plain", "application/problem+json", ""):
            with self.subTest(media_type=media_type):
                with self.assertRaisesRegex(ValueError, "Content-Type"):
                    self.probe(headers={"content-type": media_type})

    def test_rejects_unsupported_and_ambiguous_json_charsets(self):
        for media_type in (
            "application/json; charset=iso-8859-1",
            "application/json; charset=utf-16",
            "application/json; charset=\"utf-8",
            "application/json; charset=utf-8\"",
            "application/json; charset=utf-8; charset=utf-8",
            "application/json; boundary=xyz",
            "application/json;",
        ):
            with self.subTest(media_type=media_type):
                with self.assertRaisesRegex(ValueError, "Content-Type"):
                    self.probe(headers={"content-type": media_type})

    def test_accepts_declared_utf8_json_and_legacy_missing_type(self):
        for media_type in (
            "application/json",
            "application/json; charset=utf-8",
            'application/json; charset="utf-8"',
            "Application/JSON; Charset=UTF-8",
        ):
            with self.subTest(media_type=media_type):
                result, response = self.probe(headers={"content-type": media_type})
                self.assertEqual(result["port"], 8765)
                response.iter_bytes.assert_called_once_with()
        result, _ = self.probe()
        self.assertEqual(result["port"], 8765)

    def test_checks_deadline_after_each_network_chunk(self):
        # Each network read can finish before the per-read timeout. A fixed
        # 8192-byte buffer would coalesce these trickled chunks indefinitely.
        chunks = [self.PAYLOAD[:15], self.PAYLOAD[15:]]
        with self.assertRaisesRegex(ValueError, "タイムアウト"):
            self.probe(chunks=chunks, clock=[0.0, 1.0, 2.0, 3.0, 11.0])

    def test_accepts_timely_multiple_network_chunks(self):
        chunks = [self.PAYLOAD[:15], self.PAYLOAD[15:]]
        result, response = self.probe(chunks=chunks, clock=[0.0, 1.0, 2.0, 3.0, 4.0])
        self.assertEqual(result["version"], "4.56.1.1")
        response.iter_bytes.assert_called_once_with()


if __name__ == "__main__":
    unittest.main()
