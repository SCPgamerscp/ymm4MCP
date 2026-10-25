"""Connection descriptors from Windows tools may include a UTF-8 BOM (Issue #22)."""
import json
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from ymm4_connection import ConnectionConfigurationError, connection_settings


class ConnectionDescriptorBomTests(unittest.TestCase):
    def test_bom_descriptor_discovers_port_and_token(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "connection.json"
            path.write_text(
                json.dumps({"token": "file-token", "api_base": "http://127.0.0.1:9012/api"}),
                encoding="utf-8-sig",
            )
            with patch.dict(os.environ, {"YMM4_CONNECTION_FILE": str(path)}, clear=True):
                self.assertEqual(
                    connection_settings(),
                    ("http://127.0.0.1:9012/api", {"X-Ymm4-Token": "file-token"}),
                )

    def test_explicit_token_overrides_bom_descriptor_token(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "connection.json"
            path.write_text(
                json.dumps({"token": "file-token", "api_base": "http://localhost:9013/api"}),
                encoding="utf-8-sig",
            )
            with patch.dict(os.environ, {
                "YMM4_CONNECTION_FILE": str(path), "YMM4_API_TOKEN": "env-token",
            }, clear=True):
                self.assertEqual(
                    connection_settings(),
                    ("http://localhost:9013/api", {"X-Ymm4-Token": "env-token"}),
                )

    def test_invalid_utf8_descriptor_preserves_existing_fallback_policy(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "connection.json"
            path.write_bytes(b"\xff\xfeinvalid")
            with patch.dict(os.environ, {
                "YMM4_CONNECTION_FILE": str(path), "YMM4_API_TOKEN": "env-token",
            }, clear=True):
                self.assertEqual(
                    connection_settings(),
                    ("http://127.0.0.1:8765/api", {"X-Ymm4-Token": "env-token"}),
                )
            with patch.dict(os.environ, {"YMM4_CONNECTION_FILE": str(path)}, clear=True):
                with self.assertRaises(ConnectionConfigurationError):
                    connection_settings()


if __name__ == "__main__":
    unittest.main()
