"""Additional regression tests for connection descriptor fail-closed behavior (#22)."""
import json
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from ymm4_connection import (
    ConnectionConfigurationError,
    MAX_API_TOKEN_LENGTH,
    MAX_DESCRIPTOR_DEPTH,
    connection_settings,
)


class DescriptorHardeningTests(unittest.TestCase):
    def read_descriptor(self, raw: bytes, extra_env=None):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "connection.json"
            path.write_bytes(raw)
            environment = {"YMM4_CONNECTION_FILE": str(path)}
            environment.update(extra_env or {})
            with patch.dict(os.environ, environment, clear=True):
                return connection_settings()

    def test_excessive_nesting_fails_as_configuration_error(self):
        body = (b'{"token":"secret","extra":'
                + b"[" * MAX_DESCRIPTOR_DEPTH + b"0"
                + b"]" * MAX_DESCRIPTOR_DEPTH + b"}")
        with self.assertRaises(ConnectionConfigurationError):
            self.read_descriptor(body)

    def test_deep_invalid_descriptor_with_explicit_token_preserves_fallback(self):
        body = (b'{"token":"secret","extra":'
                + b"[" * (MAX_DESCRIPTOR_DEPTH + 1) + b"0"
                + b"]" * (MAX_DESCRIPTOR_DEPTH + 1) + b"}")
        self.assertEqual(
            self.read_descriptor(body, {"YMM4_API_TOKEN": "explicit"})[0],
            "http://127.0.0.1:8765/api",
        )

    def test_nesting_boundary_and_quoted_brackets_remain_valid(self):
        # Root object contributes one level to the depth budget.
        payload = json.dumps({
            "token": "secret",
            "extra": "[" * 200 + '\\"' + "}" * 200,
        }).encode("utf-8")
        self.assertEqual(self.read_descriptor(payload)[1]["X-Ymm4-Token"], "secret")
        nested = (b'{"token":"secret","extra":'
                  + b"[" * (MAX_DESCRIPTOR_DEPTH - 1) + b"0"
                  + b"]" * (MAX_DESCRIPTOR_DEPTH - 1) + b"}")
        self.assertEqual(self.read_descriptor(nested)[1]["X-Ymm4-Token"], "secret")

    def test_nonfinite_json_constants_are_rejected_in_ignored_metadata(self):
        for number in (b"NaN", b"Infinity", b"-Infinity"):
            with self.subTest(number=number):
                payload = b'{"token":"secret","extra":{"number":' + number + b"}}"
                with self.assertRaises(ConnectionConfigurationError):
                    self.read_descriptor(payload)

    def test_unpaired_unicode_surrogates_are_rejected_in_keys_and_values(self):
        samples = (
            b'{"token":"secret","extra":{"\\ud800":"bad"}}',
            b'{"token":"secret","extra":["\\udfff"]}',
        )
        for raw in samples:
            with self.subTest(raw=raw):
                with self.assertRaises(ConnectionConfigurationError):
                    self.read_descriptor(raw)

    def test_valid_unicode_surrogate_pair_remains_accepted(self):
        raw = b'{"token":"secret","extra":{"emoji":"\\ud83d\\ude00"}}'
        self.assertEqual(self.read_descriptor(raw)[1]["X-Ymm4-Token"], "secret")

    def test_explicit_empty_connection_file_cannot_fall_back_to_localappdata(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "YMM4MCP" / "connection.json"
            path.parent.mkdir()
            path.write_text('{"token":"secret"}', encoding="utf-8")
            for token in ("", "explicit"):
                with self.subTest(token=token):
                    with patch.dict(os.environ, {
                        "YMM4_CONNECTION_FILE": "",
                        "LOCALAPPDATA": folder,
                        "YMM4_API_TOKEN": token,
                    }, clear=True):
                        with self.assertRaises(ConnectionConfigurationError):
                            connection_settings()

    def test_rejects_oversized_tokens_from_environment_and_descriptor(self):
        too_long = "A" * (MAX_API_TOKEN_LENGTH + 1)
        with patch.dict(os.environ, {"YMM4_API_TOKEN": too_long}, clear=True):
            with self.assertRaises(ConnectionConfigurationError):
                connection_settings()
        with self.assertRaises(ConnectionConfigurationError):
            self.read_descriptor(json.dumps({"token": too_long}).encode("utf-8"))

    def test_token_length_boundary_and_plugin_generated_token_work(self):
        for token in ("A" * MAX_API_TOKEN_LENGTH, "0123456789abcdef" * 4):
            with self.subTest(length=len(token)):
                self.assertEqual(
                    self.read_descriptor(json.dumps({"token": token}).encode("utf-8"))
                    [1]["X-Ymm4-Token"], token,
                )


if __name__ == "__main__":
    unittest.main()
