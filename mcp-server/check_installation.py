"""Read-only Windows installation acceptance probe for ymm4MCP Issue #22.

Checks the persistent MCP install, all initialized Claude Desktop profiles, and
an authenticated YMM4 status response. Does not install, edit, or log tokens.
"""
from __future__ import annotations

import json
import os
from pathlib import Path
import sys
from urllib.parse import urlsplit

from check_connection import check as check_connection
from check_mcp_stdio import check_stdio


class InstallationError(ValueError):
    pass


MAX_CLAUDE_CONFIG_BYTES = 1024 * 1024


def _read_claude_config(path: Path):
    """Read a bounded, duplicate-safe snapshot; never log config values."""
    try:
        with path.open("rb") as stream:
            raw = stream.read(MAX_CLAUDE_CONFIG_BYTES + 1)
        if len(raw) > MAX_CLAUDE_CONFIG_BYTES:
            raise InstallationError("Claude Desktop configuration exceeds the size limit")
        config = json.loads(raw.decode("utf-8-sig"), object_pairs_hook=_unique_pairs)
    except (OSError, UnicodeError, ValueError) as exc:
        raise InstallationError(f"Invalid Claude Desktop configuration: {path.name}") from exc
    return config, raw


def _unique_pairs(pairs):
    result = {}
    for key, value in pairs:
        if key in result:
            raise InstallationError("Duplicate Claude Desktop configuration key")
        result[key] = value
    return result


def _same_path(actual, expected):
    if not isinstance(actual, str) or not actual.strip():
        return False
    return os.path.normcase(os.path.abspath(actual)) == os.path.normcase(os.path.abspath(expected))


def _config_paths(local_appdata: Path, appdata: Path):
    paths = [appdata / "Claude" / "claude_desktop_config.json"]
    packages = local_appdata / "Packages"
    if packages.is_dir():
        for directory in sorted(packages.glob("Claude_*")):
            if directory.is_dir():
                path = directory / "LocalCache" / "Roaming" / "Claude" / "claude_desktop_config.json"
                if path.parent.is_dir():
                    paths.append(path)
    return paths


def inspect_installation(local_appdata: Path, appdata: Path, *,
                         status_check=check_connection, mcp_probe=check_stdio):
    """Verify real installed paths and authenticated status without mutation."""
    root = local_appdata / "YMM4MCP" / "mcp-server"
    python = root / ".venv" / "Scripts" / "python.exe"
    server = root / "server.py"
    for path in (python, server, root / "requirements.txt"):
        if not path.is_file():
            raise InstallationError(f"Missing persistent installation file: {path.name}")
    candidates = _config_paths(local_appdata, appdata)
    configs = [path for path in candidates if path.is_file()]
    if not configs:
        raise InstallationError("No initialized Claude Desktop configuration was found")
    configured_bases = set()
    snapshots = {}
    profile_envs = []
    # Validate every initialized profile, including Microsoft Store installs.
    for path in configs:
        config, raw = _read_claude_config(path)
        snapshots[path] = (config, raw)
        if not isinstance(config, dict) or not isinstance(config.get("mcpServers"), dict):
            raise InstallationError("Claude Desktop mcpServers is missing or malformed")
        entry = config["mcpServers"].get("ymm4")
        if not isinstance(entry, dict) or not _same_path(entry.get("command"), python):
            raise InstallationError("Claude Desktop ymm4 command is not the persistent Python")
        args = entry.get("args")
        if not isinstance(args, list) or len(args) != 1 or not _same_path(args[0], server):
            raise InstallationError("Claude Desktop ymm4 script is not the persistent server")
        env = entry.get("env", {})
        if (not isinstance(env, dict) or
                any(not isinstance(key, str) or not isinstance(value, str)
                    for key, value in env.items())):
            raise InstallationError("Claude Desktop ymm4 env is malformed")
        profile_envs.append(env.copy())
        if "YMM4_API_BASE" in env:
            value = env["YMM4_API_BASE"]
            if not isinstance(value, str):
                raise InstallationError("Claude Desktop YMM4_API_BASE is invalid")
            parsed = urlsplit(value)
            try:
                port = parsed.port
            except ValueError as exc:
                raise InstallationError("Claude Desktop YMM4_API_BASE port is invalid") from exc
            if (parsed.scheme != "http" or parsed.hostname not in
                    ("localhost", "127.0.0.1", "::1") or port is None or
                    parsed.path.rstrip("/") != "/api" or parsed.query or parsed.fragment or
                    parsed.username or parsed.password):
                raise InstallationError("Claude Desktop YMM4_API_BASE must be a loopback API URL")
            configured_bases.add(value)
    if len(configured_bases) > 1:
        raise InstallationError("Claude Desktop profiles disagree on the configured port")
    # Claude may override the port while this probe runs from a plain terminal.
    original_base = os.environ.get("YMM4_API_BASE")
    try:
        if configured_bases:
            os.environ["YMM4_API_BASE"] = next(iter(configured_bases))
        else:
            # A terminal-only override is not part of Claude Desktop setup.
            # Verify the installed descriptor/default rather than an unrelated
            # process environment that can produce a false-positive status.
            os.environ.pop("YMM4_API_BASE", None)
        status = status_check()  # Includes positive AND negative token probes.
    finally:
        if original_base is None:
            os.environ.pop("YMM4_API_BASE", None)
        else:
            os.environ["YMM4_API_BASE"] = original_base
    if (not isinstance(status, dict) or status.get("status") != "running" or
            type(status.get("port")) is not int or not 1 <= status["port"] <= 65535):
        raise InstallationError("Authenticated YMM4 status check failed")
    # A healthy HTTP endpoint alone does not prove Claude can start the
    # installed MCP server. Perform a real stdio initialize + tools/list for
    # every initialized Claude profile, using that profile's environment.
    # Never invoke an editing tool or expose the profile's secrets in output.
    mcp_tool_counts = []
    for profile_env in profile_envs:
        child_env = dict(os.environ)
        child_env.pop("YMM4_API_BASE", None)
        child_env.update(profile_env)
        report = mcp_probe(python, server, env=child_env)
        if (not isinstance(report, dict) or type(report.get("tool_count")) is not int
                or report["tool_count"] < 3):
            raise InstallationError("Installed MCP server tool discovery failed")
        mcp_tool_counts.append(report["tool_count"])
    for path in configs:
        # The status and MCP probes can take seconds. A concurrent installer or Claude
        # settings edit must not turn an old validation into a success for a
        # different configuration. Re-read once, compare exact bytes, and use
        # the already validated snapshot for the final port comparison.
        config, original = snapshots[path]
        try:
            with path.open("rb") as stream:
                current = stream.read(MAX_CLAUDE_CONFIG_BYTES + 1)
        except OSError as exc:
            raise InstallationError("Claude Desktop configuration changed during acceptance check") from exc
        if current != original:
            raise InstallationError("Claude Desktop configuration changed during acceptance check")
        base = config["mcpServers"]["ymm4"].get("env", {}).get("YMM4_API_BASE")
        if base is not None and urlsplit(base).port != status["port"]:
            raise InstallationError("Claude Desktop port differs from the running YMM4 plugin")
    return {"success": True, "persistent_install": True,
            "claude_profiles_checked": len(configs),
            "authenticated_status": "running", "port": status["port"],
            "mcp_stdio": "ready", "mcp_profiles_checked": len(mcp_tool_counts),
            "mcp_tools_minimum": min(mcp_tool_counts),
            "version": status.get("version")}


def main() -> int:
    if sys.platform != "win32":
        print("FAIL: This installer acceptance probe requires Windows", file=sys.stderr)
        return 1
    if not os.environ.get("LOCALAPPDATA") or not os.environ.get("APPDATA"):
        print("FAIL: LOCALAPPDATA/APPDATA is unavailable", file=sys.stderr)
        return 1
    try:
        result = inspect_installation(Path(os.environ["LOCALAPPDATA"]),
                                      Path(os.environ["APPDATA"]))
    except Exception as exc:
        # Never include a config value, HTTP header, or token in the output.
        print(f"FAIL: {type(exc).__name__}; inspect installer and YMM4 status", file=sys.stderr)
        return 1
    print(json.dumps(result, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
