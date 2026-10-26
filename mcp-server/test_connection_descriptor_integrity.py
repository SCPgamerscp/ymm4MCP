"""Regression tests for safe YMM4 connection descriptor discovery (Issue #22)."""
import json
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from ymm4_connection import (
    ConnectionConfigurationError,
    MAX_DESCRIPTOR_BYTES,
    connection_settings,
)


class DescriptorIntegrityTests(unittest.TestCase):
    def _write(self, directory, data):
        path = Path(directory) / "connection.json"
        path.write_bytes(data)
        return str(path)

    def test_duplicate_token_or_base_is_not_silently_accepted(self):
        samples = (
            b'{"token":"first","token":"second","api_base":"http://127.0.0.1:9001/api"}',
            b'{"token":"secret","api_base":"http://127.0.0.1:9001/api",'
            b'"api_base":"http://127.0.0.1:9002/api"}',
            b'{"token":"secret","api_base":"http://127.0.0.1:9001/api",'
            b'"extra":{"name":"first","name":"second"}}',
        )
        with tempfile.TemporaryDirectory() as directory:
            for data in samples:
                with self.subTest(data=data):
                    path = self._write(directory, data)
                    with patch.dict(os.environ, {"YMM4_CONNECTION_FILE": path}, clear=True):
                        with self.assertRaises(ConnectionConfigurationError):
                            connection_settings()

    def test_duplicate_descriptor_with_explicit_token_keeps_fallback_policy(self):
        with tempfile.TemporaryDirectory() as directory:
            path = self._write(directory, b'{"token":"first","token":"second"}')
            with patch.dict(os.environ, {
                "YMM4_CONNECTION_FILE": path, "YMM4_API_TOKEN": "explicit",
            }, clear=True):
                self.assertEqual(
                    connection_settings(),
                    ("http://127.0.0.1:8765/api", {"X-Ymm4-Token": "explicit"}),
                )

    def test_oversized_descriptor_is_rejected_without_token(self):
        with tempfile.TemporaryDirectory() as directory:
            path = self._write(directory, b" " * (MAX_DESCRIPTOR_BYTES + 1))
            with patch.dict(os.environ, {"YMM4_CONNECTION_FILE": path}, clear=True):
                with self.assertRaises(ConnectionConfigurationError):
                    connection_settings()

    def test_oversized_descriptor_with_explicit_token_uses_existing_fallback(self):
        with tempfile.TemporaryDirectory() as directory:
            path = self._write(directory, b" " * (MAX_DESCRIPTOR_BYTES + 1))
            with patch.dict(os.environ, {
                "YMM4_CONNECTION_FILE": path, "YMM4_API_TOKEN": "explicit",
            }, clear=True):
                self.assertEqual(connection_settings()[0], "http://127.0.0.1:8765/api")

    def test_descriptor_at_size_limit_is_accepted(self):
        with tempfile.TemporaryDirectory() as directory:
            raw = json.dumps({
                "token": "secret", "api_base": "http://127.0.0.1:9001/api",
            }).encode("utf-8")
            path = self._write(directory, raw + b" " * (MAX_DESCRIPTOR_BYTES - len(raw)))
            with patch.dict(os.environ, {"YMM4_CONNECTION_FILE": path}, clear=True):
                self.assertEqual(connection_settings()[0], "http://127.0.0.1:9001/api")

    def test_explicit_empty_or_falsey_descriptor_base_does_not_use_default_port(self):
        with tempfile.TemporaryDirectory() as directory:
            for base in ("", None, False, 0):
                with self.subTest(base=base):
                    path = self._write(directory, json.dumps({
                        "token": "secret", "api_base": base,
                    }).encode("utf-8"))
                    with patch.dict(os.environ, {"YMM4_CONNECTION_FILE": path}, clear=True):
                        with self.assertRaises(ConnectionConfigurationError):
                            connection_settings()

    def test_explicit_empty_env_base_does_not_use_descriptor_port(self):
        with tempfile.TemporaryDirectory() as directory:
            path = self._write(directory, json.dumps({
                "token": "secret", "api_base": "http://127.0.0.1:9001/api",
            }).encode("utf-8"))
            with patch.dict(os.environ, {
                "YMM4_CONNECTION_FILE": path, "YMM4_API_BASE": "",
            }, clear=True):
                with self.assertRaises(ConnectionConfigurationError):
                    connection_settings()

    def test_bom_and_normal_descriptor_still_work(self):
        with tempfile.TemporaryDirectory() as directory:
            data = json.dumps({
                "token": "secret", "api_base": "http://localhost:9001/api",
            }).encode("utf-8")
            for prefix in (b"", b"\xef\xbb\xbf"):
                with self.subTest(prefix=prefix):
                    path = self._write(directory, prefix + data)
                    with patch.dict(os.environ, {"YMM4_CONNECTION_FILE": path}, clear=True):
                        self.assertEqual(connection_settings()[0], "http://localhost:9001/api")


if __name__ == "__main__":
    unittest.main()
