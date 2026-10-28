"""Fail closed on ambiguous, nonstandard, oversized or deeply nested status JSON."""
import unittest
from unittest.mock import MagicMock, patch

import httpx

import check_connection


class StatusJsonIntegrityTests(unittest.TestCase):
    @staticmethod
    def fake_client(response):
        client = MagicMock()
        client.__enter__.return_value = client
        wrong = MagicMock()
        wrong.__enter__.return_value = MagicMock(status_code=401)
        right = MagicMock()
        right.__enter__.return_value = response
        client.stream.side_effect = [wrong, right]
        return client

    def check_body(self, body: bytes):
        response = httpx.Response(
            200,
            content=body,
            request=httpx.Request("GET", "http://127.0.0.1:8765/api/status"),
        )
        client = self.fake_client(response)
        with patch.object(check_connection, "connection_settings", return_value=(
                "http://127.0.0.1:8765/api", {"X-Ymm4-Token": "secret"})), \
             patch.object(check_connection.httpx, "Client", return_value=client):
            return check_connection.check()

    def test_rejects_duplicate_status_fields_even_when_final_value_is_valid(self):
        for key, first in (
            ("status", '"stopped"'),
            ("port", "1234"),
            ("version", '"malicious"'),
        ):
            with self.subTest(key=key), self.assertRaisesRegex(ValueError, "重複キー"):
                self.check_body((
                    '{"status":"running","port":8765,"version":"4.56",'
                    + '"' + key + '":' + first + '}'
                ).encode("utf-8"))

    def test_rejects_nested_duplicate_keys_in_ignored_metadata(self):
        with self.assertRaisesRegex(ValueError, "重複キー"):
            self.check_body(
                b'{"status":"running","port":8765,"extra":{"a":1,"a":2}}'
            )

    def test_rejects_nonstandard_nonfinite_json_constants(self):
        for number in (b"NaN", b"Infinity", b"-Infinity"):
            with self.subTest(number=number), self.assertRaisesRegex(ValueError, "不正な数値"):
                self.check_body(
                    b'{"status":"running","port":8765,"extra":' + number + b'}'
                )

    def test_rejects_oversized_json_before_parsing(self):
        with self.assertRaisesRegex(ValueError, "大きすぎ"):
            self.check_body(
                b'{"status":"running","port":8765,"extra":"' + b"x" * 65536 + b'"}'
            )

    def test_rejects_excessive_json_nesting_as_value_error(self):
        with self.assertRaisesRegex(ValueError, "深すぎ"):
            self.check_body(
                b'{"status":"running","port":8765,"extra":'
                + b"[" * 1500 + b"0" + b"]" * 1500 + b"}"
            )

    def test_brackets_inside_quoted_json_strings_do_not_count_as_nesting(self):
        body = (
            b'{"status":"running","port":8765,"extra":"'
            + b"[" * 200 + b'\\\"' + b"{" * 200 + b'"}'
        )
        self.assertEqual(self.check_body(body)["port"], 8765)

    def test_utf16_escaped_quotes_cannot_hide_excessive_nesting(self):
        # The byte-level scanner used to mistake UTF-16 NUL bytes for escaped
        # characters, allowing deep JSON after an escaped quote.
        text = ('{"status":"running","port":8765,'
                '"extra":"escaped \\" quote","nested":'
                + "[" * 129 + "0" + "]" * 129 + "}")
        for encoding in ("utf-16", "utf-32"):
            with self.subTest(encoding=encoding), self.assertRaisesRegex(ValueError, "深すぎ"):
                self.check_body(text.encode(encoding))

    def test_accepts_utf16_utf32_and_bom_json(self):
        text = '{"status":"running","port":8765,"extra":"\\\" [[{{"}'
        for encoding in ("utf-16", "utf-32", "utf-8-sig"):
            with self.subTest(encoding=encoding):
                self.assertEqual(self.check_body(text.encode(encoding))["port"], 8765)

    def test_stream_rejects_overlimit_without_consuming_remaining_chunks(self):
        # The limit must apply while streaming, not after client.get buffers.
        chunks = iter([b"x" * 8192] * 10)
        response = MagicMock()
        response.iter_bytes.return_value = chunks
        client = self.fake_client(response)
        with patch.object(check_connection, "connection_settings", return_value=(
                "http://127.0.0.1:8765/api", {"X-Ymm4-Token": "secret"})), \
             patch.object(check_connection.httpx, "Client", return_value=client):
            with self.assertRaisesRegex(ValueError, "大きすぎ"):
                check_connection.check()
        response.iter_bytes.assert_called_once_with(chunk_size=8192)
        self.assertEqual(next(chunks), b"x" * 8192)

    def test_stream_enforces_total_response_deadline(self):
        response = MagicMock()
        response.iter_bytes.return_value = iter([b'{"status":"running","port":8765}'])
        client = self.fake_client(response)
        with patch.object(check_connection, "connection_settings", return_value=(
                "http://127.0.0.1:8765/api", {"X-Ymm4-Token": "secret"})), \
             patch.object(check_connection.httpx, "Client", return_value=client), \
             patch.object(check_connection.time, "monotonic", side_effect=[0.0, 11.0]):
            with self.assertRaisesRegex(ValueError, "タイムアウト"):
                check_connection.check()

    def test_accepts_normal_status_and_exact_size_boundary(self):
        body = b'{"status":"running","port":8765,"version":"4.56.1.1"}'
        self.assertEqual(self.check_body(body)["version"], "4.56.1.1")
        padded = body + b" " * (check_connection.MAX_STATUS_JSON_BYTES - len(body))
        self.assertEqual(len(padded), check_connection.MAX_STATUS_JSON_BYTES)
        self.assertEqual(self.check_body(padded)["port"], 8765)


if __name__ == "__main__":
    unittest.main()
