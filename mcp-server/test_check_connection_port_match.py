"""Connection checks must not accept a response from an unexpected port."""
import json
import unittest
from unittest.mock import MagicMock, patch

import check_connection


class StatusPortConsistencyTests(unittest.TestCase):
    def check_with(self, base, port):
        client = MagicMock()
        client.__enter__.return_value = client
        rejected = MagicMock(status_code=401)
        accepted = MagicMock(status_code=200)
        accepted.json.return_value = {
            "status": "running", "port": port, "success": True,
        }
        accepted.headers = {}
        accepted.iter_bytes.return_value = [json.dumps(accepted.json.return_value).encode("utf-8")]
        wrong = MagicMock()
        wrong.__enter__.return_value = rejected
        right = MagicMock()
        right.__enter__.return_value = accepted
        client.stream.side_effect = [wrong, right]
        with patch.object(check_connection, "connection_settings", return_value=(
                base, {"X-Ymm4-Token": "secret"})), \
             patch.object(check_connection.httpx, "Client", return_value=client):
            return check_connection.check()

    def test_status_port_must_match_configured_port(self):
        for base in ("http://127.0.0.1:9001/api", "http://[::1]:9001/api"):
            with self.subTest(base=base), self.assertRaisesRegex(ValueError, "一致"):
                self.check_with(base, 8765)

    def test_matching_status_port_is_accepted(self):
        self.assertEqual(self.check_with("http://localhost:9001/api", 9001)["port"], 9001)

    def test_implicit_http_port_remains_compatible(self):
        self.assertEqual(self.check_with("http://localhost/api", 8765)["port"], 8765)


if __name__ == "__main__":
    unittest.main()
