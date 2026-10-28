"""Authenticated install check validates the actual status response."""
import json
import unittest
from unittest.mock import MagicMock, patch

import check_connection


class ConnectionCheckTests(unittest.TestCase):
    def test_checks_local_status_using_discovered_token(self):
        client = MagicMock()
        client.__enter__.return_value = client
        payload = {"status": "running", "port": 8765}
        client.get.return_value.json.return_value = payload
        client.stream.return_value.__enter__.return_value.iter_bytes.return_value = [json.dumps(payload).encode()]
        with patch.object(check_connection, "connection_settings", return_value=(
                "http://127.0.0.1:8765/api", {"X-Ymm4-Token": "secret"})), \
             patch.object(check_connection.httpx, "Client", return_value=client):
            result = check_connection.check()
        client.stream.assert_called_once_with("GET", "http://127.0.0.1:8765/api/status",
                                              headers={"X-Ymm4-Token": "secret"})
        self.assertEqual(result["port"], 8765)

    def test_unexpected_status_is_failure(self):
        client = MagicMock()
        client.__enter__.return_value = client
        payload = {"status": "stopped"}
        client.get.return_value.json.return_value = payload
        client.stream.return_value.__enter__.return_value.iter_bytes.return_value = [json.dumps(payload).encode()]
        with patch.object(check_connection, "connection_settings", return_value=(
                "http://127.0.0.1:8765/api", {})), \
             patch.object(check_connection.httpx, "Client", return_value=client):
            with self.assertRaises(ValueError):
                check_connection.check()

    def test_error_bearing_or_failed_status_is_rejected(self):
        for payload in ({"status": "running", "port": 8765, "error": None},
                        {"status": "running", "port": 8765, "error_code": "FAILED"},
                        {"status": "running", "port": 8765, "success": False},
                        {"status": "running", "port": 8765, "success": 1}):
            with self.subTest(payload=payload):
                client = MagicMock()
                client.__enter__.return_value = client
                client.get.return_value.json.return_value = payload
                client.stream.return_value.__enter__.return_value.iter_bytes.return_value = [json.dumps(payload).encode()]
                with patch.object(check_connection, "connection_settings", return_value=(
                        "http://127.0.0.1:8765/api", {})), \
                     patch.object(check_connection.httpx, "Client", return_value=client):
                    with self.assertRaisesRegex(ValueError, "ステータス"):
                        check_connection.check()

    def test_invalid_status_ports_are_rejected(self):
        for port in (None, True, False, "8765", 0, -1, 65536, 1.5):
            with self.subTest(port=port):
                client = MagicMock()
                client.__enter__.return_value = client
                payload = {"status": "running", "port": port}
                client.get.return_value.json.return_value = payload
                client.stream.return_value.__enter__.return_value.iter_bytes.return_value = [json.dumps(payload).encode()]
                with patch.object(check_connection, "connection_settings", return_value=(
                        "http://127.0.0.1:8765/api", {})), \
                     patch.object(check_connection.httpx, "Client", return_value=client):
                    with self.assertRaisesRegex(ValueError, "ポート"):
                        check_connection.check()

    def test_success_true_and_boundary_ports_are_accepted(self):
        for port in (1, 65535):
            with self.subTest(port=port):
                client = MagicMock()
                client.__enter__.return_value = client
                payload = {"status": "running", "port": port, "success": True}
                client.get.return_value.json.return_value = payload
                client.stream.return_value.__enter__.return_value.iter_bytes.return_value = [json.dumps(payload).encode()]
                with patch.object(check_connection, "connection_settings", return_value=(
                        f"http://127.0.0.1:{port}/api", {})), \
                     patch.object(check_connection.httpx, "Client", return_value=client):
                    self.assertEqual(check_connection.check()["port"], port)
