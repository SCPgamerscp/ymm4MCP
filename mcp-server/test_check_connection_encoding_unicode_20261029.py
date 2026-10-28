"""Regression coverage for declared JSON encoding and Unicode/numeric integrity (#22)."""
import unittest
from unittest.mock import MagicMock, patch

import check_connection


class StatusEncodingAndUnicodeTests(unittest.TestCase):
    def probe(self, body: bytes, headers=None):
        response = MagicMock(status_code=200)
        response.headers = {} if headers is None else headers
        response.iter_bytes.return_value = [body]
        negative = MagicMock()
        negative.__enter__.return_value = MagicMock(status_code=401)
        positive = MagicMock()
        positive.__enter__.return_value = response
        client = MagicMock()
        client.__enter__.return_value = client
        client.stream.side_effect = [negative, positive]
        with patch.object(check_connection, "connection_settings", return_value=(
                "http://127.0.0.1:8765/api", {"X-Ymm4-Token": "secret"})), \
             patch.object(check_connection.httpx, "Client", return_value=client):
            return check_connection.check()

    def test_declared_utf8_rejects_utf16_payload(self):
        payload = '{"status":"running","port":8765}'.encode("utf-16")
        with self.assertRaisesRegex(ValueError, "文字コード"):
            self.probe(payload, {"content-type": "application/json; charset=utf-8"})

    def test_declared_json_rejects_utf32_payload(self):
        payload = '{"status":"running","port":8765}'.encode("utf-32")
        with self.assertRaisesRegex(ValueError, "文字コード"):
            self.probe(payload, {"content-type": "application/json"})

    def test_declared_utf8_accepts_utf8_bom(self):
        payload = '{"status":"running","port":8765}'.encode("utf-8-sig")
        self.assertEqual(
            self.probe(payload, {"content-type": "application/json; charset=utf-8"})["port"],
            8765,
        )

    def test_legacy_missing_content_type_accepts_utf16(self):
        payload = '{"status":"running","port":8765}'.encode("utf-16")
        self.assertEqual(self.probe(payload)["port"], 8765)

    def test_rejects_unpaired_surrogate_in_ignored_nested_value(self):
        payload = b'{"status":"running","port":8765,"extra":{"label":"\\ud800"}}'
        with self.assertRaisesRegex(ValueError, "Unicode"):
            self.probe(payload)

    def test_rejects_unpaired_surrogate_in_ignored_key(self):
        payload = b'{"status":"running","port":8765,"\\udfff":"ignored"}'
        with self.assertRaisesRegex(ValueError, "Unicode"):
            self.probe(payload)

    def test_accepts_valid_surrogate_pair_in_optional_metadata(self):
        payload = b'{"status":"running","port":8765,"emoji":"\\ud83d\\ude00"}'
        self.assertEqual(self.probe(payload)["port"], 8765)

    def test_rejects_oversized_float_token_in_ignored_metadata(self):
        payload = b'{"status":"running","port":8765,"extra":0.' + b"0" * 127 + b"}"
        with self.assertRaisesRegex(ValueError, "小数が長すぎ"):
            self.probe(payload)

    def test_accepts_float_token_at_128_character_boundary(self):
        payload = b'{"status":"running","port":8765,"extra":0.' + b"0" * 126 + b"}"
        self.assertEqual(self.probe(payload)["port"], 8765)


if __name__ == "__main__":
    unittest.main()
