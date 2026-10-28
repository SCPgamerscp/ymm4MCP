"""The install connection check must not trust unprintable host version data."""
import json
import unittest
from unittest.mock import MagicMock, patch

import check_connection


class StatusVersionIntegrityTests(unittest.TestCase):
    def check_payload(self, payload):
        client = MagicMock()
        client.__enter__.return_value = client
        rejected = MagicMock(status_code=401)
        accepted = MagicMock(status_code=200)
        accepted.json.return_value = payload
        accepted.iter_bytes.return_value = [json.dumps(payload).encode("utf-8")]
        client.stream.return_value.__enter__.return_value = accepted
        # Standalone #244 makes only the positive request; when #228 is
        # composed first, the negative authentication probe must also work.
        # Select by token, not call order, to test both compositions.
        client.get.side_effect = lambda _url, headers: (
            accepted if headers.get("X-Ymm4-Token") == "secret" else rejected
        )
        with patch.object(check_connection, "connection_settings", return_value=(
                "http://127.0.0.1:8765/api", {"X-Ymm4-Token": "secret"})), \
             patch.object(check_connection.httpx, "Client", return_value=client):
            return check_connection.check()

    def test_rejects_control_characters_and_direction_overrides(self):
        for version in ("4.56\\nFAKE", "4.56\\x1b[31m", "4.56\\u202e", "4.56\\u2066"):
            version = version.encode("ascii", errors="backslashreplace").decode("unicode_escape")
            with self.subTest(version=repr(version)), self.assertRaisesRegex(ValueError, "バージョン"):
                self.check_payload({"status": "running", "port": 8765, "version": version})

    def test_rejects_invalid_version_shapes_and_oversized_strings(self):
        for version in (None, True, 123, {}, [], "", "v" * 129):
            with self.subTest(version=version), self.assertRaisesRegex(ValueError, "バージョン"):
                self.check_payload({"status": "running", "port": 8765, "version": version})

    def test_preserves_compatible_version_and_optional_absence(self):
        for payload, expected in (
            ({"status": "running", "port": 8765, "version": "4.56.1.1"}, "4.56.1.1"),
            ({"status": "running", "port": 8765}, None),
            ({"status": "running", "port": 8765, "version": "v" * 128}, "v" * 128),
        ):
            with self.subTest(payload=payload):
                self.assertEqual(self.check_payload(payload)["version"], expected)


if __name__ == "__main__":
    unittest.main()
