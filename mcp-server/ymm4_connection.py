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
    path = os.environ.get("YMM4_CONNECTION_FILE")
    if not path:
        local = os.environ.get("LOCALAPPDATA")
        if local:
            path = str(Path(local) / "YMM4MCP" / "connection.json")
    if path:
        try:
            with open(path, encoding="utf-8") as source:
                descriptor = json.load(source)
            if not isinstance(descriptor, dict):
                raise ValueError("object required")
        except (OSError, ValueError) as exc:
            if not token:
                raise ConnectionConfigurationError(
                    "YMM4接続情報を読めません。プラグインを起動し、YMM4_CONNECTION_FILEを確認してください"
                ) from exc
            descriptor = {}
    elif not token:
        raise ConnectionConfigurationError(
            "YMM4_CONNECTION_FILE または YMM4_API_TOKEN を設定してください"
        )
    if not token:
        token = descriptor.get("token", "")
    base = os.environ.get("YMM4_API_BASE") or descriptor.get("api_base") or "http://127.0.0.1:8765/api"
    if not isinstance(base, str):
        raise ConnectionConfigurationError("YMM4_API_BASE は文字列で指定してください")
    # urllib.parse strips tabs/newlines and leading C0 controls silently.
    # Reject them before parsing so the validated URL is the URL we use.
    if any(c.isspace() or ord(c) < 32 or ord(c) == 127 for c in base):
        raise ConnectionConfigurationError("YMM4_API_BASE に空白や制御文字は使用できません")
    if "?" in base or "#" in base:
        raise ConnectionConfigurationError("YMM4_API_BASE にクエリやフラグメントは使用できません")
    try:
        parsed = urlsplit(base)
    except ValueError as exc:
        raise ConnectionConfigurationError("YMM4_API_BASE のURL形式が不正です") from exc
    if (parsed.scheme != "http" or "@" in parsed.netloc
            or parsed.hostname not in {"127.0.0.1", "localhost", "::1"}
            or parsed.username or parsed.password or parsed.query or parsed.fragment
            or parsed.path.rstrip("/") != "/api"):
        raise ConnectionConfigurationError("YMM4_API_BASE はローカルHTTPの /api URLのみ指定できます")
    try:
        if parsed.port is not None and not 1024 <= parsed.port <= 65535:
            raise ValueError("port out of range")
    except ValueError as exc:
        raise ConnectionConfigurationError("ポートは1024〜65535で指定してください") from exc
    if not isinstance(token, str) or not token or any(not (33 <= ord(c) <= 126) for c in token):
        raise ConnectionConfigurationError("有効なYMM4 APIトークンがありません")
    return base.rstrip("/"), {"X-Ymm4-Token": token}


def advanced_enabled() -> bool:
    return os.environ.get("YMM4_ENABLE_ADVANCED", "").lower() in {"1", "true"}
