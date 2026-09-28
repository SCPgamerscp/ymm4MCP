"""Authenticated install check validates the actual status response."""
import unittest
from unittest.mock import MagicMock, patch

import check_connection


class ConnectionCheckTests(unittest.TestCase):
    def test_checks_local_status_using_discovered_token(self):
        client = MagicMock()
        client.__enter__.return_value = client
        client.get.return_value.json.return_value = {"status": "running", "port": 8765}
        with patch.object(check_connection, "connection_settings", return_value=(
                "http://127.0.0.1:8765/api", {"X-Ymm4-Token": "secret"})), \
             patch.object(check_connection.httpx, "Client", return_value=client):
            result = check_connection.check()
        client.get.assert_called_once_with("http://127.0.0.1:8765/api/status",
                                           headers={"X-Ymm4-Token": "secret"})
        self.assertEqual(result["port"], 8765)

    def test_unexpected_status_is_failure(self):
        client = MagicMock()
        client.__enter__.return_value = client
        client.get.return_value.json.return_value = {"status": "stopped"}
        with patch.object(check_connection, "connection_settings", return_value=(
                "http://127.0.0.1:8765/api", {})), \
             patch.object(check_connection.httpx, "Client", return_value=client):
            with self.assertRaises(ValueError):
                check_connection.check()
