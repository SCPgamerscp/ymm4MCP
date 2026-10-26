"""Check the installed MCP server's authenticated local YMM4 connection."""
import sys
from urllib.parse import urlsplit

import httpx

from ymm4_connection import ConnectionConfigurationError, connection_settings


def check() -> dict:
    base, headers = connection_settings()
    # This endpoint is strictly loopback. Environment HTTP(S)_PROXY settings
    # must never receive the plugin's X-Ymm4-Token credential.
    with httpx.Client(timeout=5.0, trust_env=False) as client:
        response = client.get(base + "/status", headers=headers)
        response.raise_for_status()
        status = response.json()
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
    return {"status": status["status"], "version": status.get("version"), "port": port}


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
