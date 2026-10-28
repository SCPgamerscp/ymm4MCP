"""Local authenticated connection checks must ignore ambient proxy variables (#22)."""
import os
import unittest
from unittest.mock import MagicMock, call, patch

import check_connection


class LocalConnectionProxyIsolationTests(unittest.TestCase):
    def test_local_token_is_not_sent_via_environment_proxy(self):
        for base in (
            "http://127.0.0.1:8765/api",
            "http://localhost:8765/api",
            "http://[::1]:8765/api",
        ):
            with self.subTest(base=base):
                client = MagicMock()
                client.__enter__.return_value = client
                rejected = MagicMock(status_code=401)
                accepted = MagicMock(status_code=200)
                accepted.json.return_value = {
                    "status": "running", "port": 8765, "success": True
                }
                client.get.side_effect = [rejected, accepted]
                proxy_environment = {
                    "HTTP_PROXY": "http://127.0.0.1:9999",
                    "HTTPS_PROXY": "http://127.0.0.1:9999",
                    "ALL_PROXY": "http://127.0.0.1:9999",
                    "NO_PROXY": "",
                }
                with patch.dict(os.environ, proxy_environment), \
                     patch.object(check_connection, "connection_settings",
                                  return_value=(base, {"X-Ymm4-Token": "secret"})), \
                     patch.object(check_connection.httpx, "Client",
                                  return_value=client) as client_factory:
                    result = check_connection.check()

                client_factory.assert_called_once_with(timeout=5.0, trust_env=False)
                self.assertEqual(client.get.call_args_list, [
                    call(base + "/status", headers={"X-Ymm4-Token": "0ecret"}),
                    call(base + "/status", headers={"X-Ymm4-Token": "secret"}),
                ])
                self.assertEqual(result["port"], 8765)


if __name__ == "__main__":
    unittest.main()
