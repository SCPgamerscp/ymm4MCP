"""Read the plugin's per-user connection descriptor without caching rotated tokens."""
import json
import os
from pathlib import Path
from urllib.parse import urlsplit


class ConnectionConfigurationError(ValueError):
    pass


# The plugin writes a small JSON descriptor. A corrupted or attacker-controlled
# descriptor must not cause unbounded reads during every MCP request.
MAX_DESCRIPTOR_BYTES = 64 * 1024
MAX_DESCRIPTOR_DEPTH = 128
MAX_API_TOKEN_LENGTH = 256


def _unique_json_object(pairs):
    result = {}
    for key, value in pairs:
        if key in result:
            raise ValueError(f"duplicate connection descriptor key: {key}")
        result[key] = value
    return result


def _reject_nonfinite_constant(value: str):
    """Connection metadata must use standards-compliant JSON numbers."""
    raise ValueError("nonfinite connection descriptor number")


def _descriptor_too_deep(body: str) -> bool:
    """Bound JSON nesting without relying on the interpreter recursion limit."""
    depth = 0
    quoted = False
    escaped = False
    for char in body:
        if quoted:
            if escaped:
                escaped = False
            elif char == "\\":
                escaped = True
            elif char == '"':
                quoted = False
        elif char == '"':
            quoted = True
        elif char in "[{":
            depth += 1
            if depth > MAX_DESCRIPTOR_DEPTH:
                return True
        elif char in "]}":
            depth -= 1
    return False


def _reject_unpaired_surrogates(value: object) -> None:
    """Invalid escaped Unicode must not survive in ignored descriptor metadata."""
    pending = [value]
    while pending:
        item = pending.pop()
        if isinstance(item, str):
            if any(0xD800 <= ord(char) <= 0xDFFF for char in item):
                raise ValueError("unpaired Unicode surrogate in connection descriptor")
        elif isinstance(item, dict):
            pending.extend(item.keys())
            pending.extend(item.values())
        elif isinstance(item, list):
            pending.extend(item)


def connection_settings() -> tuple[str, dict[str, str]]:
    token = os.environ.get("YMM4_API_TOKEN", "")
    descriptor = {}
    # An explicitly empty path is a configuration error, not a request to
    # silently discover a different YMM4 instance via LOCALAPPDATA.
    if "YMM4_CONNECTION_FILE" in os.environ and not os.environ["YMM4_CONNECTION_FILE"]:
        raise ConnectionConfigurationError("YMM4_CONNECTION_FILE は空にできません")
    path = os.environ.get("YMM4_CONNECTION_FILE")
    if not path:
        local = os.environ.get("LOCALAPPDATA")
        if local:
            path = str(Path(local) / "YMM4MCP" / "connection.json")
    if path:
        try:
            with open(path, "rb") as source:
                raw = source.read(MAX_DESCRIPTOR_BYTES + 1)
            if len(raw) > MAX_DESCRIPTOR_BYTES:
                raise ValueError("connection descriptor is too large")
            decoded = raw.decode("utf-8-sig")
            if _descriptor_too_deep(decoded):
                raise ValueError("connection descriptor is too deeply nested")
            descriptor = json.loads(
                decoded,
                object_pairs_hook=_unique_json_object,
                parse_constant=_reject_nonfinite_constant,
            )
            _reject_unpaired_surrogates(descriptor)
            if not isinstance(descriptor, dict):
                raise ValueError("object required")
        except (OSError, ValueError, RecursionError) as exc:
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
    # A present but malformed base is a configuration error, not a signal to
    # silently connect to the default port (possibly another YMM4 instance).
    if "YMM4_API_BASE" in os.environ:
        base = os.environ["YMM4_API_BASE"]
    elif "api_base" in descriptor:
        base = descriptor["api_base"]
    else:
        base = "http://127.0.0.1:8765/api"
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
        # YMM4 never binds its API on the implicit HTTP port 80. An omitted
        # port can otherwise silently connect to a different local service.
        if parsed.port is None or not 1024 <= parsed.port <= 65535:
            raise ValueError("port missing or out of range")
    except ValueError as exc:
        raise ConnectionConfigurationError("ポートは1024〜65535で指定してください") from exc
    if (not isinstance(token, str) or not token or len(token) > MAX_API_TOKEN_LENGTH
            or any(not (33 <= ord(c) <= 126) for c in token)):
        raise ConnectionConfigurationError("有効なYMM4 APIトークンがありません")
    return base.rstrip("/"), {"X-Ymm4-Token": token}


def advanced_enabled() -> bool:
    return os.environ.get("YMM4_ENABLE_ADVANCED", "").lower() in {"1", "true"}
