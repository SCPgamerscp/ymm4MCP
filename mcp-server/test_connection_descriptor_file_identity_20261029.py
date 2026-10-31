"""Regression tests for descriptor file identity and canonical API URLs (#22)."""
import json
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from ymm4_connection import ConnectionConfigurationError, connection_settings


class DescriptorFileTypeAndUrlTests(unittest.TestCase):
    def read(self, path, token=None):
        env = {"YMM4_CONNECTION_FILE": str(path)}
        if token is not None:
            env["YMM4_API_TOKEN"] = token
        with patch.dict(os.environ, env, clear=True):
            return connection_settings()

    def test_regular_descriptor_file_remains_supported(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "connection.json"
            path.write_text(json.dumps({"token": "secret"}), encoding="utf-8")
            self.assertEqual(self.read(path)[1]["X-Ymm4-Token"], "secret")

    def test_directory_is_not_a_connection_descriptor(self):
        with tempfile.TemporaryDirectory() as folder:
            with self.assertRaises(ConnectionConfigurationError):
                self.read(folder)

    def test_descriptor_symlink_is_rejected(self):
        with tempfile.TemporaryDirectory() as folder:
            target = Path(folder) / "real.json"
            target.write_text('{"token":"secret"}', encoding="utf-8")
            link = Path(folder) / "alias.json"
            try:
                link.symlink_to(target)
            except (OSError, NotImplementedError) as exc:
                self.skipTest(f"symlinks unavailable: {exc}")
            with self.assertRaises(ConnectionConfigurationError):
                self.read(link)

    @unittest.skipUnless(hasattr(os, "mkfifo"), "FIFO unsupported on this platform")
    def test_fifo_is_rejected_without_attempting_a_blocking_read(self):
        with tempfile.TemporaryDirectory() as folder:
            fifo = Path(folder) / "connection.fifo"
            os.mkfifo(fifo)
            with self.assertRaises(ConnectionConfigurationError):
                self.read(fifo)

    def test_explicit_token_preserves_unreadable_descriptor_fallback(self):
        with tempfile.TemporaryDirectory() as folder:
            self.assertEqual(
                self.read(folder, token="explicit"),
                ("http://127.0.0.1:8765/api", {"X-Ymm4-Token": "explicit"}),
            )

    def test_rejects_ambiguous_api_path_with_repeated_trailing_slashes(self):
        for path in ("/api//", "/api///", "/api/.", "/api/%2F"):
            with self.subTest(path=path):
                with patch.dict(os.environ, {
                    "YMM4_API_TOKEN": "secret",
                    "YMM4_API_BASE": "http://127.0.0.1:8765" + path,
                }, clear=True):
                    with self.assertRaisesRegex(ConnectionConfigurationError, "URL"):
                        connection_settings()

    def test_accepts_exact_api_path_and_one_optional_trailing_slash(self):
        for path in ("/api", "/api/"):
            with self.subTest(path=path):
                with patch.dict(os.environ, {
                    "YMM4_API_TOKEN": "secret",
                    "YMM4_API_BASE": "http://127.0.0.1:8765" + path,
                }, clear=True):
                    self.assertEqual(connection_settings()[0], "http://127.0.0.1:8765/api")


if __name__ == "__main__":
    unittest.main()
