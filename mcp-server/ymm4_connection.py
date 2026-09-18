"""Read the plugin's per-user connection descriptor without caching rotated tokens."""
import json
import os
from pathlib import Path
from urllib.parse import urlsplit


class ConnectionConfigurationError(ValueError):
    pass


def connection_settings() -> tuple[str, dict[str, str]]:
    token = os.environ.get("YMM4_API_TOKEN", "")
    descriptor = {}
    if not token:
        path = os.environ.get("YMM4_CONNECTION_FILE")
        if not path:
            local = os.environ.get("LOCALAPPDATA")
            if not local:
                raise ConnectionConfigurationError(
                    "YMM4_CONNECTION_FILE または YMM4_API_TOKEN を設定してください"
                )
            path = str(Path(local) / "YMM4MCP" / "connection.json")
        try:
            with open(path, encoding="utf-8") as source:
                descriptor = json.load(source)
            if not isinstance(descriptor, dict):
                raise ValueError("object required")
            token = descriptor.get("token", "")
        except (OSError, ValueError) as exc:
            raise ConnectionConfigurationError(
                "YMM4接続情報を読めません。プラグインを起動し、YMM4_CONNECTION_FILEを確認してください"
            ) from exc
    base = os.environ.get("YMM4_API_BASE") or descriptor.get("api_base") or "http://127.0.0.1:8765/api"
    if not isinstance(base, str):
        raise ConnectionConfigurationError("YMM4_API_BASE は文字列で指定してください")
    parsed = urlsplit(base)
    if (parsed.scheme != "http" or parsed.hostname not in {"127.0.0.1", "localhost", "::1"}
            or parsed.username or parsed.password or parsed.query or parsed.fragment
            or parsed.path.rstrip("/") != "/api"):
        raise ConnectionConfigurationError("YMM4_API_BASE はローカルHTTPの /api URLのみ指定できます")
    try:
        if parsed.port is not None and not 1024 <= parsed.port <= 65535:
            raise ValueError("port out of range")
    except ValueError as exc:
        raise ConnectionConfigurationError("ポートは1024〜65535で指定してください") from exc
    if not isinstance(token, str) or not token or not token.isascii() or any(c.isspace() for c in token):
        raise ConnectionConfigurationError("有効なYMM4 APIトークンがありません")
    return base.rstrip("/"), {"X-Ymm4-Token": token}


def advanced_enabled() -> bool:
    return os.environ.get("YMM4_ENABLE_ADVANCED", "").lower() in {"1", "true"}
