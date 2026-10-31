"""Regression tests for bounded status JSON numbers and unambiguous HTTP framing (#22)."""
import unittest
from unittest.mock import MagicMock, patch

import check_connection


class StatusNumericAndTransferIntegrityTests(unittest.TestCase):
    def probe(self, payload=None, headers=None):
        response = MagicMock(status_code=200)
        response.headers = {} if headers is None else headers
        response.iter_bytes.return_value = [
            payload if payload is not None
            else b'{"status":"running","port":8765,"version":"4.56.1.1"}'
        ]
        client = MagicMock()
        client.__enter__.return_value = client
        negative = MagicMock()
        negative.__enter__.return_value = MagicMock(status_code=401)
        positive = MagicMock()
        positive.__enter__.return_value = response
        client.stream.side_effect = [negative, positive]
        with patch.object(check_connection, "connection_settings", return_value=(
                "http://127.0.0.1:8765/api", {"X-Ymm4-Token": "secret"})), \
             patch.object(check_connection.httpx, "Client", return_value=client):
            return check_connection.check()

    def test_rejects_exponent_overflow_even_in_ignored_fields(self):
        for token in (b"1e309", b"-1e9999", b"1.2e10000"):
            with self.subTest(token=token), self.assertRaisesRegex(ValueError, "不正な数値"):
                self.probe(b'{"status":"running","port":8765,"extra":' + token + b'}')

    def test_rejects_nested_float_overflow(self):
        with self.assertRaisesRegex(ValueError, "不正な数値"):
            self.probe(b'{"status":"running","port":8765,"extra":[{"n":1e400}]}')

    def test_accepts_finite_floats_in_optional_metadata(self):
        result = self.probe(
            b'{"status":"running","port":8765,"extra":[1e308,-1e-308,0.25]}'
        )
        self.assertEqual(result["port"], 8765)

    def test_rejects_overlong_integer_even_in_ignored_fields(self):
        with self.assertRaisesRegex(ValueError, "整数が長すぎ"):
            self.probe(
                b'{"status":"running","port":8765,"extra":'
                + b"9" * 129 + b"}"
            )

    def test_accepts_bounded_integer_in_optional_metadata(self):
        self.assertEqual(self.probe(
            b'{"status":"running","port":8765,"extra":'
            + b"9" * 128 + b"}"
        )["port"], 8765)

    def test_rejects_transfer_encoding_and_content_length_combination(self):
        for transfer in ("chunked", "CHUNKED", " chunked "):
            with self.subTest(transfer=transfer), self.assertRaisesRegex(ValueError, "フレーミング"):
                self.probe(headers={
                    "transfer-encoding": transfer,
                    "content-length": "51",
                })

    def test_rejects_unsupported_or_ambiguous_transfer_encodings(self):
        for transfer in ("gzip", "gzip, chunked", "chunked, chunked",
                         "chunked, gzip", "", "identity"):
            with self.subTest(transfer=transfer), self.assertRaisesRegex(ValueError, "Transfer-Encoding"):
                self.probe(headers={"transfer-encoding": transfer})

    def test_accepts_normal_chunked_status_response(self):
        self.assertEqual(
            self.probe(headers={"transfer-encoding": "chunked"})["port"], 8765
        )


if __name__ == "__main__":
    unittest.main()
