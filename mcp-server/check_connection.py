"""Check the installed MCP server's authenticated local YMM4 connection."""
import sys

import httpx

from ymm4_connection import ConnectionConfigurationError, connection_settings


def check() -> dict:
    base, headers = connection_settings()
    with httpx.Client(timeout=5.0) as client:
        response = client.get(base + "/status", headers=headers)
        response.raise_for_status()
        status = response.json()
    if not isinstance(status, dict) or status.get("status") != "running":
        raise ValueError("YMM4のステータスを確認できません")
    return {"status": status["status"], "version": status.get("version"), "port": status.get("port")}


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
