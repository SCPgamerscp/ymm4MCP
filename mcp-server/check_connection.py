"""Check the installed MCP server's authenticated local YMM4 connection."""
import sys
from urllib.parse import urlsplit

import httpx

from ymm4_connection import ConnectionConfigurationError, connection_settings


MAX_STATUS_JSON_BYTES = 65536
MAX_STATUS_JSON_DEPTH = 128


def _unique_json_object(pairs: list[tuple[str, object]]) -> dict:
    """Do not let a later duplicate status field silently override evidence."""
    result = {}
    for key, value in pairs:
        if key in result:
            raise ValueError("YMM4のステータスJSONに重複キーがあります")
        result[key] = value
    return result


def _reject_nonfinite_json_constant(value: str):
    """JSON does not permit NaN/Infinity, even in ignored fields."""
    raise ValueError("YMM4のステータスJSONに不正な数値があります")


def _status_json_too_deep(body: bytes) -> bool:
    """Bound structural nesting without counting brackets inside JSON strings."""
    depth = 0
    quoted = False
    escaped = False
    for byte in body:
        if quoted:
            if escaped:
                escaped = False
            elif byte == 92:  # backslash
                escaped = True
            elif byte == 34:  # double quote
                quoted = False
        elif byte == 34:
            quoted = True
        elif byte in (91, 123):  # [ or {
            depth += 1
            if depth > MAX_STATUS_JSON_DEPTH:
                return True
        elif byte in (93, 125):  # ] or }
            depth -= 1
    return False


def check() -> dict:
    base, headers = connection_settings()
    # This endpoint is strictly loopback. Environment HTTP(S)_PROXY settings
    # must never receive the plugin's X-Ymm4-Token credential.
    with httpx.Client(timeout=5.0, trust_env=False) as client:
        response = client.get(base + "/status", headers=headers)
        response.raise_for_status()
        # The local endpoint has a tiny status payload. Bound JSON parsing
        # even when an unexpected service responds with an oversized body.
        if len(response.content) > MAX_STATUS_JSON_BYTES:
            raise ValueError("YMM4のステータスJSONが大きすぎます")
        if _status_json_too_deep(response.content):
            raise ValueError("YMM4のステータスJSONが深すぎます")
        try:
            status = response.json(
                object_pairs_hook=_unique_json_object,
                parse_constant=_reject_nonfinite_json_constant,
            )
        except RecursionError as exc:
            raise ValueError("YMM4のステータスJSONが深すぎます") from exc
    if (not isinstance(status, dict) or status.get("status") != "running"
            or "error" in status or "error_code" in status
            or ("success" in status and status["success"] is not True)):
        raise ValueError("YMM4のステータスを確認できません")
    port = status.get("port")
    if isinstance(port, bool) or not isinstance(port, int) or not 1 <= port <= 65535:
        raise ValueError("YMM4の応答ポートが不正です")
    # A successful HTTP response on another port does not prove that the
    # configured plugin instance is the one we intended to contact.
    configured_port = urlsplit(base).port
    if configured_port is not None and port != configured_port:
        raise ValueError("YMM4の応答ポートが接続先ポートと一致しません")
    version = status.get("version")
    # Version is printed in the installer terminal. Never display control
    # sequences, Unicode direction overrides, or arbitrarily large values
    # returned by an unexpected service on the configured loopback port.
    if "version" in status and (
            not isinstance(version, str) or not version or len(version) > 128
            or not version.isprintable()):
        raise ValueError("YMM4の応答バージョンが不正です")
    return {"status": status["status"], "version": version, "port": port}


def main() -> int:
    try:
        status = check()
    except (ConnectionConfigurationError, httpx.HTTPError, ValueError) as exc:
        print(f"接続確認に失敗しました: {exc}", file=sys.stderr)
        print("YMM4を起動し、MCPプラグインの自動起動とポート設定を確認してください。", file=sys.stderr)
        return 1
    print(f"YMM4 MCP 接続成功: port={status['port']} version={status['version']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
