"""Regression tests for descriptor replacement and mid-read mutation (#22)."""
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from ymm4_connection import ConnectionConfigurationError, connection_settings


class DescriptorReadRaceTests(unittest.TestCase):
    def _settings(self, path, token=None):
        env = {"YMM4_CONNECTION_FILE": str(path)}
        if token is not None:
            env["YMM4_API_TOKEN"] = token
        with patch.dict(os.environ, env, clear=True):
            return connection_settings()

    def test_unchanged_regular_file_still_works(self):
        with tempfile.TemporaryDirectory() as folder:
            descriptor = Path(folder) / "connection.json"
            descriptor.write_text('{"token":"unchanged"}', encoding="utf-8")
            self.assertEqual(
                self._settings(descriptor)[1]["X-Ymm4-Token"], "unchanged"
            )

    def test_replaced_regular_file_between_stat_and_open_is_rejected(self):
        with tempfile.TemporaryDirectory() as folder:
            descriptor = Path(folder) / "connection.json"
            replacement = Path(folder) / "replacement.json"
            descriptor.write_text('{"token":"old"}', encoding="utf-8")
            replacement.write_text('{"token":"new"}', encoding="utf-8")
            real_open = os.open

            def replace_then_open(path, flags, *args, **kwargs):
                if os.fspath(path) == str(descriptor):
                    os.replace(replacement, descriptor)
                return real_open(path, flags, *args, **kwargs)

            with patch("ymm4_connection.os.open", side_effect=replace_then_open):
                with self.assertRaises(ConnectionConfigurationError):
                    self._settings(descriptor)

    def test_file_modified_during_read_is_rejected(self):
        with tempfile.TemporaryDirectory() as folder:
            descriptor = Path(folder) / "connection.json"
            descriptor.write_text('{"token":"before"}', encoding="utf-8")
            real_fdopen = os.fdopen

            class MutatingReader:
                def __init__(self, stream):
                    self.stream = stream

                def __enter__(self):
                    self.stream.__enter__()
                    return self

                def __exit__(self, *args):
                    return self.stream.__exit__(*args)

                def fileno(self):
                    return self.stream.fileno()

                def read(self, size):
                    data = self.stream.read(size)
                    descriptor.write_text(
                        '{"token":"modified-during-read"}', encoding="utf-8"
                    )
                    return data

            def mutating_fdopen(fd, mode):
                return MutatingReader(real_fdopen(fd, mode))

            with patch("ymm4_connection.os.fdopen", side_effect=mutating_fdopen):
                with self.assertRaises(ConnectionConfigurationError):
                    self._settings(descriptor)

    def test_explicit_token_keeps_documented_fallback_on_race(self):
        with tempfile.TemporaryDirectory() as folder:
            descriptor = Path(folder) / "connection.json"
            replacement = Path(folder) / "replacement.json"
            descriptor.write_text('{"token":"old"}', encoding="utf-8")
            replacement.write_text('{"token":"new"}', encoding="utf-8")
            real_open = os.open

            def replace_then_open(path, flags, *args, **kwargs):
                if os.fspath(path) == str(descriptor):
                    os.replace(replacement, descriptor)
                return real_open(path, flags, *args, **kwargs)

            with patch("ymm4_connection.os.open", side_effect=replace_then_open):
                self.assertEqual(
                    self._settings(descriptor, token="explicit"),
                    ("http://127.0.0.1:8765/api",
                     {"X-Ymm4-Token": "explicit"}),
                )


if __name__ == "__main__":
    unittest.main()
