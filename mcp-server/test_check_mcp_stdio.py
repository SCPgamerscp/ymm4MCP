"""MCP stdio acceptance: real SDK subprocess, no YMM4 host required."""
import os
from pathlib import Path
import sys
import tempfile
import unittest

from check_mcp_stdio import McpStdioProbeError, check_stdio


class McpStdioAcceptanceTests(unittest.TestCase):
    def test_installed_server_initializes_and_exposes_tools(self):
        result = check_stdio(
            Path(sys.executable), Path(__file__).with_name("server.py"),
            env=dict(os.environ), timeout_seconds=25.0)
        self.assertGreaterEqual(result["tool_count"], 4)
        self.assertNotIn("token", repr(result).lower())

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
