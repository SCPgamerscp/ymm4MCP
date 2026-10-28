"""Authenticated install checks require both negative and positive probes."""
import json
import unittest
from unittest.mock import MagicMock, call, patch

import check_connection


def fake_client(payload, rejected_status=401):
    client = MagicMock()
    client.__enter__.return_value = client
    rejected = MagicMock(status_code=rejected_status)
    accepted = MagicMock(status_code=200)
    accepted.json.return_value = payload
    accepted.headers = {}
    accepted.iter_bytes.return_value = [json.dumps(payload).encode("utf-8")]
    wrong = MagicMock()
    wrong.__enter__.return_value = rejected
    right = MagicMock()
    right.__enter__.return_value = accepted
    client.stream.side_effect = [wrong, right]
    return client


class ConnectionCheckTests(unittest.TestCase):
    def run_check(self, payload, token="secret", rejected_status=401, base=None):
        client = fake_client(payload, rejected_status)
        base = base or "http://127.0.0.1:8765/api"
        with patch.object(check_connection, "connection_settings", return_value=(
                base, {"X-Ymm4-Token": token})), \
             patch.object(check_connection.httpx, "Client", return_value=client):
            result = check_connection.check()
        return result, client

    def test_checks_local_status_using_discovered_token(self):
        result, client = self.run_check({"status": "running", "port": 8765})
        self.assertEqual(client.stream.call_args_list, [
            call("GET", "http://127.0.0.1:8765/api/status",
                 headers={"X-Ymm4-Token": "0ecret"}),
            call("GET", "http://127.0.0.1:8765/api/status",
                 headers={"X-Ymm4-Token": "secret", "Accept-Encoding": "identity"}),
        ])
        self.assertEqual(result["port"], 8765)

    def test_authentication_must_reject_wrong_same_length_token(self):
        for rejected_status in (200, 202, 403, 404, 500):
            with self.subTest(rejected_status=rejected_status):
                client = fake_client({"status": "running", "port": 8765},
                                     rejected_status)
                with patch.object(check_connection, "connection_settings", return_value=(
                        "http://127.0.0.1:8765/api",
                        {"X-Ymm4-Token": "secret"})), \
                     patch.object(check_connection.httpx, "Client", return_value=client):
                    with self.assertRaisesRegex(ValueError, "認証"):
                        check_connection.check()
                self.assertEqual(client.stream.call_count, 1,
                                 "Do not send a valid token if the negative probe fails")

    def test_missing_token_fails_before_network(self):
        for headers in ({}, {"X-Ymm4-Token": ""}, {"X-Ymm4-Token": None}):
            with self.subTest(headers=headers), \
                 patch.object(check_connection, "connection_settings", return_value=(
                     "http://127.0.0.1:8765/api", headers)), \
                 patch.object(check_connection.httpx, "Client") as factory:
                with self.assertRaisesRegex(ValueError, "トークン"):
                    check_connection.check()
                factory.assert_not_called()

    def test_unexpected_status_is_failure(self):
        with self.assertRaisesRegex(ValueError, "ステータス"):
            self.run_check({"status": "stopped"})

    def test_error_bearing_or_failed_status_is_rejected(self):
        for payload in ({"status": "running", "port": 8765, "error": None},
                        {"status": "running", "port": 8765, "error_code": "FAILED"},
                        {"status": "running", "port": 8765, "success": False},
                        {"status": "running", "port": 8765, "success": 1}):
            with self.subTest(payload=payload), self.assertRaisesRegex(ValueError, "ステータス"):
                self.run_check(payload)

    def test_invalid_status_ports_are_rejected(self):
        for port in (None, True, False, "8765", 0, -1, 65536, 1.5):
            with self.subTest(port=port), self.assertRaisesRegex(ValueError, "ポート"):
                self.run_check({"status": "running", "port": port})

    def test_success_true_and_boundary_ports_are_accepted(self):
        for port in (1, 65535):
            with self.subTest(port=port):
                result, _ = self.run_check(
                    {"status": "running", "port": port, "success": True},
                    base=f"http://127.0.0.1:{port}/api")
                self.assertEqual(result["port"], port)


if __name__ == "__main__":
    unittest.main()
