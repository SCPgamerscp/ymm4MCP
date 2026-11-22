"""MCP stdio acceptance: real SDK subprocess, no YMM4 host required."""
from http.server import BaseHTTPRequestHandler, HTTPServer
import json
import os
from pathlib import Path
import sys
import tempfile
import threading
import unittest

from check_mcp_stdio import McpStdioProbeError, check_stdio


class McpStdioAcceptanceTests(unittest.TestCase):
    def test_installed_server_initializes_and_exposes_tools(self):
        result = check_stdio(
            Path(sys.executable), Path(__file__).with_name("server.py"),
            env=dict(os.environ), timeout_seconds=25.0)
        self.assertGreaterEqual(result["tool_count"], 3)
        self.assertNotIn("token", repr(result).lower())

    def start_status_server(self, *, report_wrong_port=False):
        requests = []

        class StatusHandler(BaseHTTPRequestHandler):
            def do_GET(self):
                requests.append(self.path)
                if self.path != "/api/status" or self.headers.get("X-Ymm4-Token") != "fixture-token":
                    self.send_error(401)
                    return
                port = self.server.server_port
                if report_wrong_port:
                    port = port + 1 if port < 65535 else port - 1
                report = json.dumps({"status": "running", "port": port}).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.send_header("Content-Length", str(len(report)))
                self.end_headers()
                self.wfile.write(report)

            def log_message(self, *_args):
                pass

        server = HTTPServer(("127.0.0.1", 0), StatusHandler)
        self.addCleanup(server.server_close)
        thread = threading.Thread(target=server.serve_forever, daemon=True)
        thread.start()
        self.addCleanup(thread.join)
        self.addCleanup(server.shutdown)
        env = dict(os.environ)
        env.pop("YMM4_CONNECTION_FILE", None)
        env.update(YMM4_API_BASE=f"http://127.0.0.1:{server.server_port}/api",
                   YMM4_API_TOKEN="fixture-token")
        return server.server_port, env, requests

    def test_installed_server_reads_status_through_real_stdio(self):
        port, env, requests = self.start_status_server()
        result = check_stdio(Path(sys.executable), Path(__file__).with_name("server.py"),
                             env=env, expected_port=port, timeout_seconds=25.0)
        self.assertEqual(result["status_port"], port)
        self.assertEqual(requests, ["/api/status"])

    def test_installed_server_rejects_wrong_live_status_port(self):
        port, env, requests = self.start_status_server(report_wrong_port=True)
        with self.assertRaisesRegex(McpStdioProbeError, "expected port"):
            check_stdio(Path(sys.executable), Path(__file__).with_name("server.py"),
                        env=env, expected_port=port, timeout_seconds=25.0)
        self.assertEqual(requests, ["/api/status"])

    def test_missing_server_rejected(self):
        with self.assertRaises(McpStdioProbeError):
            check_stdio(Path(sys.executable), Path("missing-ymm4-server.py"),
                        env=dict(os.environ))

    def test_invalid_child_environment_rejected(self):
        with self.assertRaises(McpStdioProbeError):
            check_stdio(Path(sys.executable), Path(__file__).with_name("server.py"),
                        env={"YMM4_API_BASE": 123})

    def test_non_mcp_child_cannot_pass(self):
        with tempfile.TemporaryDirectory() as directory:
            script = Path(directory) / "not_mcp.py"
            script.write_text("print('not an MCP server', flush=True)\n",
                              encoding="utf-8")
            with self.assertRaises(McpStdioProbeError):
                check_stdio(Path(sys.executable), script,
                            env=dict(os.environ), timeout_seconds=5.0)


if __name__ == "__main__":
    unittest.main()
