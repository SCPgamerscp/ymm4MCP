"""Check the installed MCP server's authenticated local YMM4 connection."""
import sys
from urllib.parse import urlsplit

import httpx

from ymm4_connection import ConnectionConfigurationError, connection_settings


def check() -> dict:
    base, headers = connection_settings()
    token = headers.get("X-Ymm4-Token")
    if not isinstance(token, str) or not token:
        raise ValueError("認証付き接続の確認に必要なYMM4 APIトークンがありません")

    # A successful status response alone does not prove authentication.
    # Probe with a WRONG token of the SAME length to catch length-only checks.
    invalid_token = ("0" if token[0] != "0" else "1") + token[1:]
    invalid_headers = {**headers, "X-Ymm4-Token": invalid_token}
    url = base + "/status"

    # Both probes stay on loopback; environment proxies must not see tokens.
    with httpx.Client(timeout=5.0, trust_env=False) as client:
        rejected = client.get(url, headers=invalid_headers)
        if rejected.status_code != 401:
            raise ValueError(
                "YMM4の認証が有効であることを確認できません"
                "（不正なトークンへの応答がHTTP 401ではありません）"
            )
        response = client.get(url, headers=headers)
        response.raise_for_status()
        status = response.json()
    if (not isinstance(status, dict) or status.get("status") != "running"
            or "error" in status or "error_code" in status
            or ("success" in status and status["success"] is not True)):
        raise ValueError("YMM4のステータスを確認できません")
    port = status.get("port")
    if isinstance(port, bool) or not isinstance(port, int) or not 1 <= port <= 65535:
        raise ValueError("YMM4の応答ポートが不正です")
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
