"""Read-only installation acceptance tests; no real YMM4 required."""
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import Mock, patch

from check_installation import InstallationError, inspect_installation


class InstallationAcceptanceTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.local = self.root / "Local"
        self.appdata = self.root / "Roaming"
        self.server = self.local / "YMM4MCP" / "mcp-server"
        self.python = self.server / ".venv" / "Scripts" / "python.exe"
        for file in (self.python, self.server / "server.py", self.server / "requirements.txt"):
            file.parent.mkdir(parents=True, exist_ok=True)
            file.write_text("test")
        self.classic = self.appdata / "Claude" / "claude_desktop_config.json"
        self.store = (self.local / "Packages" / "Claude_Test" / "LocalCache"
                      / "Roaming" / "Claude" / "claude_desktop_config.json")
        self.status = Mock(return_value={"status": "running", "port": 8765,
                                         "version": "1.5.0"})
        self.write_config(self.classic)

    def write_config(self, path, *, command=None, args=None, env=None):
        path.parent.mkdir(parents=True, exist_ok=True)
        entry = {"command": str(self.python) if command is None else command,
                 "args": [str(self.server / "server.py")] if args is None else args}
        if env is not None:
            entry["env"] = env
        path.write_text(json.dumps({"mcpServers": {"ymm4": entry}}), encoding="utf-8")

    def inspect(self):
        return inspect_installation(self.local, self.appdata, status_check=self.status)

    def test_real_install_and_authenticated_status(self):
        self.write_config(self.store, env={"YMM4_API_BASE": "http://127.0.0.1:8765/api"})
        report = self.inspect()
        self.assertTrue(report["success"])
        self.assertEqual(report["claude_profiles_checked"], 2)
        self.status.assert_called_once()
        self.assertNotIn("token", json.dumps(report).lower())

    def test_missing_persistent_server_refuses_to_check_status(self):
        (self.server / "server.py").unlink()
        with self.assertRaisesRegex(InstallationError, "Missing persistent"):
            self.inspect()
        self.status.assert_not_called()

    def test_stale_store_profile_fails_even_if_classic_is_valid(self):
        self.write_config(self.store, command="C:/old/removed/python.exe")
        with self.assertRaisesRegex(InstallationError, "persistent Python"):
            self.inspect()
        self.status.assert_not_called()

    def test_wrong_script_path_fails(self):
        self.write_config(self.classic, args=["C:/temporary/server.py"])
        with self.assertRaisesRegex(InstallationError, "persistent server"):
            self.inspect()

    def test_configured_port_is_used_for_live_probe_and_restored(self):
        self.write_config(self.classic, env={"YMM4_API_BASE": "http://127.0.0.1:8765/api"})
        with patch.dict("os.environ", {"YMM4_API_BASE": "http://127.0.0.1:9999/api"}):
            self.status.side_effect = lambda: (
                self.assertEqual(__import__("os").environ["YMM4_API_BASE"],
                                 "http://127.0.0.1:8765/api") or
                {"status": "running", "port": 8765})
            self.inspect()
            self.assertEqual(__import__("os").environ["YMM4_API_BASE"],
                             "http://127.0.0.1:9999/api")

    def test_conflicting_claude_profiles_fail_before_network(self):
        self.write_config(self.classic, env={"YMM4_API_BASE": "http://127.0.0.1:8765/api"})
        self.write_config(self.store, env={"YMM4_API_BASE": "http://127.0.0.1:8766/api"})
        with self.assertRaisesRegex(InstallationError, "disagree"):
            self.inspect()
        self.status.assert_not_called()

    def test_port_mismatch_fails(self):
        self.write_config(self.classic, env={"YMM4_API_BASE": "http://127.0.0.1:9000/api"})
        with self.assertRaisesRegex(InstallationError, "port differs"):
            self.inspect()

    def test_external_api_url_rejected_before_network(self):
        self.write_config(self.classic, env={"YMM4_API_BASE": "http://example.org:8765/api"})
        with self.assertRaisesRegex(InstallationError, "loopback"):
            self.inspect()
        self.status.assert_not_called()

    def test_duplicate_json_keys_rejected(self):
        self.classic.write_text('{"mcpServers":{"ymm4":{},"ymm4":{}}}')
        with self.assertRaisesRegex(InstallationError, "Invalid Claude Desktop"):
            self.inspect()

    def test_missing_claude_config_fails(self):
        self.classic.unlink()
        with self.assertRaisesRegex(InstallationError, "No initialized"):
            self.inspect()

    def test_invalid_authenticated_status_fails(self):
        self.status.return_value = {"status": "stopped", "port": 8765}
        with self.assertRaisesRegex(InstallationError, "Authenticated"):
            self.inspect()


if __name__ == "__main__":
    unittest.main()
