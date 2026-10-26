"""The local YMM4 API never uses HTTP's implicit port 80 (Issue #22)."""
import json
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from ymm4_connection import ConnectionConfigurationError, connection_settings


class ExplicitPluginPortTests(unittest.TestCase):
    def test_rejects_missing_port_in_explicit_base(self):
        for base in ("http://127.0.0.1/api", "http://localhost/api",
                     "http://[::1]/api", "http://localhost:/api"):
            with self.subTest(base=base):
                with patch.dict(os.environ, {
                    "YMM4_API_TOKEN": "test-token", "YMM4_API_BASE": base,
                }, clear=True):
                    with self.assertRaisesRegex(ConnectionConfigurationError, "ポート"):
                        connection_settings()

    def test_rejects_missing_port_from_connection_descriptor(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "connection.json"
            path.write_text(json.dumps({
                "token": "descriptor-token", "api_base": "http://localhost/api",
            }), encoding="utf-8")
            with patch.dict(os.environ, {"YMM4_CONNECTION_FILE": str(path)}, clear=True):
                with self.assertRaisesRegex(ConnectionConfigurationError, "ポート"):
                    connection_settings()

    def test_explicit_valid_local_ports_remain_accepted(self):
        for base in ("http://127.0.0.1:1024/api",
                     "http://localhost:8765/api",
                     "http://[::1]:65535/api"):
            with self.subTest(base=base):
                with patch.dict(os.environ, {
                    "YMM4_API_TOKEN": "test-token", "YMM4_API_BASE": base,
                }, clear=True):
                    self.assertEqual(connection_settings()[0], base)

    def test_default_port_is_explicit_and_unchanged(self):
        with patch.dict(os.environ, {"YMM4_API_TOKEN": "test-token"}, clear=True):
            self.assertEqual(connection_settings()[0], "http://127.0.0.1:8765/api")


if __name__ == "__main__":
    unittest.main()
