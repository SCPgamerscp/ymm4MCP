"""Connection checks must not accept a response from an unexpected port."""
import json
import unittest
from unittest.mock import MagicMock, patch

import check_connection


class StatusPortConsistencyTests(unittest.TestCase):
    def check_with(self, base, port):
        client = MagicMock()
        client.__enter__.return_value = client
        payload = {"status": "running", "port": port, "success": True}
        client.get.return_value.json.return_value = payload
        client.stream.return_value.__enter__.return_value.iter_bytes.return_value = [json.dumps(payload).encode()]
        with patch.object(check_connection, "connection_settings", return_value=(base, {})), \
             patch.object(check_connection.httpx, "Client", return_value=client):
            return check_connection.check()

    def test_status_port_must_match_configured_port(self):
        for base in ("http://127.0.0.1:9001/api", "http://[::1]:9001/api"):
            with self.subTest(base=base), self.assertRaisesRegex(ValueError, "一致"):
                self.check_with(base, 8765)

    def test_matching_status_port_is_accepted(self):
        self.assertEqual(self.check_with("http://localhost:9001/api", 9001)["port"], 9001)

    def test_implicit_http_port_remains_compatible(self):
        # If no explicit port is configured, the status payload remains authoritative.
        self.assertEqual(self.check_with("http://localhost/api", 8765)["port"], 8765)


if __name__ == "__main__":
    unittest.main()
