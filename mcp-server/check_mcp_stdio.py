"""Real, read-only MCP stdio handshake for installed ymm4MCP (#22).

Starts the installed server exactly as Claude Desktop would, negotiates the
MCP protocol and discovers tools. It never calls a tool or logs credentials.
"""
from __future__ import annotations

import os
from pathlib import Path

import anyio
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

REQUIRED_TOOLS = frozenset({
    "ymm4_interact", "ymm4_preview", "ymm4_advanced", "ymm4_analyze_video"
})


class McpStdioProbeError(RuntimeError):
    pass


def check_stdio(python: Path, server: Path, *, env: dict[str, str],
                timeout_seconds: float = 20.0) -> dict:
    """Verify real initialize + tools/list, without executing editing tools."""
    if (not python.is_file() or not server.is_file() or
            not isinstance(env, dict) or
            any(not isinstance(k, str) or not isinstance(v, str)
                for k, v in env.items()) or
            not 0 < timeout_seconds <= 60):
        raise McpStdioProbeError("MCP stdio probe inputs are invalid")

    async def run() -> dict:
        params = StdioServerParameters(
            command=str(python), args=[str(server)], env=env)
        # Child stderr may contain local paths or secrets; never relay it.
        with open(os.devnull, "w", encoding="utf-8") as errlog:
            with anyio.fail_after(timeout_seconds):
                async with stdio_client(params, errlog=errlog) as (read, write):
                    async with ClientSession(read, write) as session:
                        initialized = await session.initialize()
                        info = getattr(initialized, "serverInfo", None)
                        if getattr(info, "name", None) != "ymm4-mcp":
                            raise McpStdioProbeError("Unexpected MCP server identity")
                        listed = await session.list_tools()
                        names = {tool.name for tool in listed.tools}
                        if not REQUIRED_TOOLS.issubset(names):
                            raise McpStdioProbeError("Required MCP tools unavailable")
                        return {"tool_count": len(names)}

    try:
        return anyio.run(run)
    except Exception:
        # Exception text from SDK/process may contain a profile secret.
        raise McpStdioProbeError(
            "Installed MCP server did not complete stdio initialization "
            "and tool discovery") from None
