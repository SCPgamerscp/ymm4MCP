"""Read-only Windows installation acceptance probe for ymm4MCP Issue #22.

Checks the persistent MCP install, all initialized Claude Desktop profiles, and
an authenticated YMM4 status response. Does not install, edit, or log tokens.
"""
from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
import stat
import sys
from urllib.parse import urlsplit

from check_connection import check as check_connection
from check_mcp_stdio import check_stdio


class InstallationError(ValueError):
    pass


MAX_CLAUDE_CONFIG_BYTES = 1024 * 1024


def _read_claude_config(path: Path):
    """Read a bounded, duplicate-safe snapshot; never log config values."""
    if not path.is_file():
        raise InstallationError("Missing Claude Desktop configuration in an initialized profile")
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
    if not isinstance(actual, str) or not Path(actual).is_absolute():
        return False
    return os.path.normcase(os.path.abspath(actual)) == os.path.normcase(os.path.abspath(expected))


def _installation_file_signature(path: Path) -> tuple[int, int, int, int, int]:
    try:
        info = path.stat()
    except OSError as exc:
        raise InstallationError(f"Missing persistent installation file: {path.name}") from exc
    if not stat.S_ISREG(info.st_mode):
        raise InstallationError(f"Missing persistent installation file: {path.name}")
    # Ignore access time because running the probe can update it.
    return info.st_dev, info.st_ino, info.st_size, info.st_mtime_ns, info.st_ctime_ns


def _initialized_config_paths(local_appdata: Path, appdata: Path):
    classic = appdata / "Claude"
    paths = [classic / "claude_desktop_config.json"] if classic.is_dir() else []
    packages = local_appdata / "Packages"
    if packages.is_dir():
        for directory in sorted(packages.glob("Claude_*")):
            if directory.is_dir():
                path = directory / "LocalCache" / "Roaming" / "Claude" / "claude_desktop_config.json"
                if path.parent.is_dir():
                    paths.append(path)
    # Profile initialization is indicated by the configuration directory.
    # A missing config in such a profile must fail, not disappear from the check.
    return paths


def _authenticated_status_for_profile(status_check, profile_env):
    """Probe using exactly this Claude profile's YMM4 connection overrides.

    In particular, do not inherit a terminal-only token, descriptor or port:
    an unrelated healthy connection must not make a broken profile pass.
    Restore the caller's environment even if authentication fails.
    """
    keys = ("YMM4_API_BASE", "YMM4_API_TOKEN", "YMM4_CONNECTION_FILE")
    original = {key: os.environ.get(key) for key in keys}
    try:
        for key in keys:
            if key in profile_env:
                os.environ[key] = profile_env[key]
            else:
                os.environ.pop(key, None)
        return status_check()  # Includes positive AND negative token probes.
    finally:
        for key, value in original.items():
            if value is None:
                os.environ.pop(key, None)
            else:
                os.environ[key] = value


def inspect_installation(local_appdata: Path, appdata: Path, *,
                         expected_port: int | None = None,
                         status_check=check_connection, mcp_probe=check_stdio):
    """Verify real installed paths and authenticated status without mutation."""
    if expected_port is not None and (type(expected_port) is not int or not 1024 <= expected_port <= 65535):
        raise InstallationError("Expected YMM4 port must be an integer in 1024..65535")
    root = local_appdata / "YMM4MCP" / "mcp-server"
    python = root / ".venv" / "Scripts" / "python.exe"
    server = root / "server.py"
    installed_files = {path: _installation_file_signature(path)
                       for path in (python, server, root / "requirements.txt")}
    configs = _initialized_config_paths(local_appdata, appdata)
    if not configs:
        raise InstallationError("No initialized Claude Desktop configuration was found")
    configured_ports = set()
    snapshots = {}
    profile_envs = []
    # Validate every initialized profile, including Microsoft Store installs.
    for path in configs:
        config, raw = _read_claude_config(path)
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
        port = None
        if "YMM4_API_BASE" in env:
            value = env["YMM4_API_BASE"]
            if not isinstance(value, str):
                raise InstallationError("Claude Desktop YMM4_API_BASE is invalid")
            try:
                parsed = urlsplit(value)
                port = parsed.port
            except ValueError as exc:
                raise InstallationError("Claude Desktop YMM4_API_BASE port is invalid") from exc
            if (parsed.scheme != "http" or parsed.hostname not in
                    ("localhost", "127.0.0.1", "::1") or port is None or
                    parsed.path.rstrip("/") != "/api" or parsed.query or parsed.fragment or
                    parsed.username or parsed.password):
                raise InstallationError("Claude Desktop YMM4_API_BASE must be a loopback API URL")
            configured_ports.add(port)
        snapshots[path] = (raw, port)
    if len(configured_ports) > 1:
        raise InstallationError("Claude Desktop profiles disagree on the configured port")
    if expected_port is not None and configured_ports - {expected_port}:
        raise InstallationError("Claude Desktop configured port differs from the expected port")
    # An MCP stdio handshake only proves the child starts. It does not prove
    # that each Claude profile can authenticate to the YMM4 HTTP endpoint.
    # Check every profile with its own token, descriptor and base URL, rather
    # than accepting one healthy profile as evidence for all of them.
    status = None
    for profile_env in profile_envs:
        current = _authenticated_status_for_profile(status_check, profile_env)
        if (not isinstance(current, dict) or current.get("status") != "running" or
                type(current.get("port")) is not int or
                not 1 <= current["port"] <= 65535):
            raise InstallationError("Authenticated YMM4 status check failed")
        if status is not None and current["port"] != status["port"]:
            raise InstallationError("Claude Desktop profiles disagree on the running port")
        if expected_port is not None and current["port"] != expected_port:
            raise InstallationError("Running YMM4 port differs from the expected port")
        status = current
    # Verify the installed process can read YMM4 status with each profile's
    # actual environment. A successful handshake alone only lists tools.
    mcp_tool_counts = []
    for profile_env in profile_envs:
        child_env = dict(os.environ)
        for key in ("YMM4_API_BASE", "YMM4_API_TOKEN", "YMM4_CONNECTION_FILE"):
            child_env.pop(key, None)
        child_env.update(profile_env)
        report = mcp_probe(python, server, env=child_env, expected_port=status["port"])
        if (not isinstance(report, dict) or type(report.get("tool_count")) is not int
                or report["tool_count"] < 3):
            raise InstallationError("Installed MCP server tool discovery failed")
        if type(report.get("status_port")) is not int or report["status_port"] != status["port"]:
            raise InstallationError("Installed MCP server YMM4 status port was not verified")
        mcp_tool_counts.append(report["tool_count"])
    if expected_port is not None:
        # Explicit profile overrides could otherwise hide a stale plugin descriptor.
        descriptor = root.parent / "connection.json"
        current = _authenticated_status_for_profile(status_check, {"YMM4_CONNECTION_FILE": str(descriptor)})
        if not isinstance(current, dict) or current.get("status") != "running" or \
                type(current.get("port")) is not int or current["port"] != expected_port:
            raise InstallationError("Default YMM4 connection descriptor differs from the expected port")
    for path in configs:
        # The status and MCP probes can take seconds. A concurrent installer or Claude
        # settings edit must not turn an old validation into a success for a
        # different configuration. Re-read once, compare exact bytes, and use
        # the already validated snapshot for the final port comparison.
        original, configured_port = snapshots[path]
        try:
            with path.open("rb") as stream:
                current = stream.read(MAX_CLAUDE_CONFIG_BYTES + 1)
        except OSError as exc:
            raise InstallationError("Claude Desktop configuration changed during acceptance check") from exc
        if current != original:
            raise InstallationError("Claude Desktop configuration changed during acceptance check")
        if configured_port is not None and configured_port != status["port"]:
            raise InstallationError("Claude Desktop port differs from the running YMM4 plugin")
    if set(_initialized_config_paths(local_appdata, appdata)) != snapshots.keys():
        raise InstallationError("Claude Desktop profiles changed during acceptance check")
    for path, original in installed_files.items():
        try:
            current = _installation_file_signature(path)
        except InstallationError as exc:
            raise InstallationError(f"Installation files changed during acceptance check: {path.name}") from exc
        if current != original:
            raise InstallationError(f"Installation files changed during acceptance check: {path.name}")
    return {"success": True, "persistent_install": True,
            "claude_profiles_checked": len(configs),
            "authenticated_status": "running", "port": status["port"],
            "mcp_stdio": "ready", "mcp_status": "running",
            "mcp_profiles_checked": len(mcp_tool_counts),
            "mcp_tools_minimum": min(mcp_tool_counts),
            "version": status.get("version")}


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--expected-port", type=int,
                        help="verify that profiles and the default plugin connection use this port")
    args = parser.parse_args(argv)
    if args.expected_port is not None and not 1024 <= args.expected_port <= 65535:
        parser.error("--expected-port must be in 1024..65535")
    if sys.platform != "win32":
        print("FAIL: This installer acceptance probe requires Windows", file=sys.stderr)
        return 1
    if not os.environ.get("LOCALAPPDATA") or not os.environ.get("APPDATA"):
        print("FAIL: LOCALAPPDATA/APPDATA is unavailable", file=sys.stderr)
        return 1
    try:
        result = inspect_installation(Path(os.environ["LOCALAPPDATA"]),
                                      Path(os.environ["APPDATA"]), expected_port=args.expected_port)
    except InstallationError as exc:
        print(f"FAIL: {exc}", file=sys.stderr)
        return 1
    except Exception as exc:
        # Never include a config value, HTTP header, or token in the output.
        print(f"FAIL: {type(exc).__name__}; inspect installer and YMM4 status", file=sys.stderr)
        return 1
    print(json.dumps(result, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
