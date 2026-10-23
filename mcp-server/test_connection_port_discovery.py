"""Regression coverage for connection descriptor discovery and validation (Issue #22)."""
import json
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from ymm4_connection import ConnectionConfigurationError, connection_settings


class ConnectionPortDiscoveryTests(unittest.TestCase):
    def descriptor(self, directory, token="file-token", base="http://127.0.0.1:9001/api"):
        path = Path(directory) / "connection.json"
        path.write_text(json.dumps({"token": token, "api_base": base}), encoding="utf-8")
        return str(path)

    def test_environment_token_still_discovers_descriptor_port(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = self.descriptor(tmp)
            with patch.dict(os.environ, {"YMM4_API_TOKEN": "env-token", "YMM4_CONNECTION_FILE": path}, clear=True):
                self.assertEqual(connection_settings(), ("http://127.0.0.1:9001/api", {"X-Ymm4-Token": "env-token"}))

    def test_explicit_base_takes_priority(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = self.descriptor(tmp)
            with patch.dict(os.environ, {"YMM4_API_TOKEN": "env-token", "YMM4_API_BASE": "http://localhost:9010/api", "YMM4_CONNECTION_FILE": path}, clear=True):
                self.assertEqual(connection_settings()[0], "http://localhost:9010/api")

    def test_descriptor_token_rotation_is_not_cached(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = self.descriptor(tmp, token="first")
            with patch.dict(os.environ, {"YMM4_CONNECTION_FILE": path}, clear=True):
                self.assertEqual(connection_settings()[1]["X-Ymm4-Token"], "first")
                self.descriptor(tmp, token="second", base="http://127.0.0.1:9002/api")
                self.assertEqual(connection_settings(), ("http://127.0.0.1:9002/api", {"X-Ymm4-Token": "second"}))

    def test_unreadable_descriptor_falls_back_only_with_explicit_token(self):
        with tempfile.TemporaryDirectory() as tmp:
            missing = str(Path(tmp) / "missing.json")
            with patch.dict(os.environ, {"YMM4_API_TOKEN": "env-token", "YMM4_CONNECTION_FILE": missing}, clear=True):
                self.assertEqual(connection_settings()[0], "http://127.0.0.1:8765/api")
            with patch.dict(os.environ, {"YMM4_CONNECTION_FILE": missing}, clear=True):
                with self.assertRaises(ConnectionConfigurationError):
                    connection_settings()

    def test_invalid_descriptor_base_is_not_silently_ignored(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = self.descriptor(tmp, base="http://example.com/api")
            with patch.dict(os.environ, {"YMM4_API_TOKEN": "env-token", "YMM4_CONNECTION_FILE": path}, clear=True):
                with self.assertRaises(ConnectionConfigurationError):
                    connection_settings()

    def test_invalid_descriptor_json_with_explicit_token_falls_back(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "connection.json"
            path.write_text("{broken", encoding="utf-8")
            with patch.dict(os.environ, {"YMM4_API_TOKEN": "env-token", "YMM4_CONNECTION_FILE": str(path)}, clear=True):
                self.assertEqual(connection_settings()[0], "http://127.0.0.1:8765/api")

    def test_reject_url_whitespace_controls_and_empty_delimiters(self):
        values = (" http://127.0.0.1:9001/api", "\nhttp://127.0.0.1:9001/api",
                  "http://127.0.0.1:9001/api\t", "http://127.0.0.1:9001/api?",
                  "http://127.0.0.1:9001/api#", "http://@127.0.0.1:9001/api")
        for base in values:
            with self.subTest(base=repr(base)):
                with patch.dict(os.environ, {"YMM4_API_TOKEN": "token", "YMM4_API_BASE": base}, clear=True):
                    with self.assertRaises(ConnectionConfigurationError):
                        connection_settings()

    def test_malformed_ipv6_is_configuration_error(self):
        for base in ("http://[::1/api", "http://[::1]bad/api"):
            with self.subTest(base=base):
                with patch.dict(os.environ, {"YMM4_API_TOKEN": "token", "YMM4_API_BASE": base}, clear=True):
                    with self.assertRaises(ConnectionConfigurationError):
                        connection_settings()

    def test_reject_invalid_ports(self):
        for port in ("0", "1023", "65536", "abc"):
            with self.subTest(port=port):
                with patch.dict(os.environ, {"YMM4_API_TOKEN": "token", "YMM4_API_BASE": f"http://127.0.0.1:{port}/api"}, clear=True):
                    with self.assertRaises(ConnectionConfigurationError):
                        connection_settings()

    def test_reject_control_characters_in_explicit_token(self):
        # NUL cannot be stored in os.environ; the descriptor-token test covers it.
        for token in ("a b", "a\x1bb", "a\x7fb", "a\nb"):
            with self.subTest(token=repr(token)):
                with patch.dict(os.environ, {"YMM4_API_TOKEN": token}, clear=True):
                    with self.assertRaises(ConnectionConfigurationError):
                        connection_settings()

    def test_reject_control_characters_in_descriptor_token(self):
        with tempfile.TemporaryDirectory() as tmp:
            for token in ("a\x00b", "a\x1bb", "a\x7fb"):
                with self.subTest(token=repr(token)):
                    path = self.descriptor(tmp, token=token)
                    with patch.dict(os.environ, {"YMM4_CONNECTION_FILE": path}, clear=True):
                        with self.assertRaises(ConnectionConfigurationError):
                            connection_settings()

    def test_localappdata_default_descriptor_with_env_token(self):
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp) / "YMM4MCP"
            folder.mkdir()
            self.descriptor(folder)
            with patch.dict(os.environ, {"LOCALAPPDATA": tmp, "YMM4_API_TOKEN": "env-token"}, clear=True):
                self.assertEqual(connection_settings()[0], "http://127.0.0.1:9001/api")


if __name__ == "__main__":
    unittest.main()
