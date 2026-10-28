"""Authenticated status checks must bound wire response framing (#22)."""
import unittest
from unittest.mock import MagicMock, patch

import check_connection


class StatusResponseFramingTests(unittest.TestCase):
    PAYLOAD = b'{"status":"running","port":8765,"version":"4.56.1.1"}'

    def probe(self, headers=None, chunks=None):
        response = MagicMock(status_code=200)
        response.headers = {} if headers is None else headers
        response.iter_bytes.return_value = [self.PAYLOAD] if chunks is None else chunks
        client = MagicMock()
        client.__enter__.return_value = client
        wrong = MagicMock()
        wrong.__enter__.return_value = MagicMock(status_code=401)
        right = MagicMock()
        right.__enter__.return_value = response
        client.stream.side_effect = [wrong, right]
        with patch.object(check_connection, "connection_settings", return_value=(
                "http://127.0.0.1:8765/api", {"X-Ymm4-Token": "secret"})), \
             patch.object(check_connection.httpx, "Client", return_value=client):
            try:
                return check_connection.check()
            finally:
                positive = client.stream.call_args_list[1]
                self.assertEqual(positive.kwargs["headers"]["Accept-Encoding"], "identity")

    def test_rejects_compressed_responses_before_decoding(self):
        for encoding in ("gzip", "deflate", "br", "gzip, identity"):
            with self.subTest(encoding=encoding), self.assertRaisesRegex(ValueError, "圧縮"):
                self.probe(headers={"content-encoding": encoding})

    def test_rejects_oversized_declared_length_without_buffering(self):
        with self.assertRaisesRegex(ValueError, "大きすぎ"):
            self.probe(headers={"content-length": "65537"}, chunks=iter(()))

    def test_rejects_ambiguous_or_malformed_declared_length(self):
        for length in ("-1", "5, 5", " 55", "５５", "000000", "nan"):
            with self.subTest(length=length), self.assertRaisesRegex(ValueError, "Content-Length"):
                self.probe(headers={"content-length": length})

    def test_rejects_truncated_or_contradictory_response_length(self):
        for length in (0, len(self.PAYLOAD) - 1, len(self.PAYLOAD) + 1):
            with self.subTest(length=length), self.assertRaisesRegex(ValueError, "一致"):
                self.probe(headers={"content-length": str(length)})

    def test_accepts_exact_length_and_streamed_without_length(self):
        self.assertEqual(
            self.probe(headers={"content-length": str(len(self.PAYLOAD))})["port"], 8765
        )
        self.assertEqual(
            self.probe(headers={"transfer-encoding": "chunked"},
                       chunks=[self.PAYLOAD[:10], self.PAYLOAD[10:]])["port"], 8765
        )
        self.assertEqual(
            self.probe(headers={"content-encoding": "identity"})["port"], 8765
        )


if __name__ == "__main__":
    unittest.main()
