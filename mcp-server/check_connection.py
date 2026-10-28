"""Check the installed MCP server's authenticated local YMM4 connection."""
import json
import math
import sys
import time
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


def _finite_json_float(value: str) -> float:
    """Reject overflowed JSON exponents, not just nonstandard NaN tokens."""
    number = float(value)
    if not math.isfinite(number):
        raise ValueError("YMM4のステータスJSONに不正な数値があります")
    return number


def _bounded_json_int(value: str) -> int:
    """Bound integer parsing even when Python's global digit guard is disabled."""
    if len(value.lstrip("-")) > 128:
        raise ValueError("YMM4のステータスJSONの整数が長すぎます")
    return int(value)


def _status_json_too_deep(body: str) -> bool:
    """Bound structural nesting in decoded JSON, including UTF-16/32 inputs."""
    depth = 0
    quoted = False
    escaped = False
    for char in body:
        if quoted:
            if escaped:
                escaped = False
            elif char == '\\':  # backslash
                escaped = True
            elif char == '"':  # double quote
                quoted = False
        elif char == '"':
            quoted = True
        elif char in ('[', '{'):
            depth += 1
            if depth > MAX_STATUS_JSON_DEPTH:
                return True
        elif char in (']', '}'):
            depth -= 1
    return False


def check() -> dict:
    base, headers = connection_settings()
    token = headers.get("X-Ymm4-Token")
    if not isinstance(token, str) or not token:
        raise ValueError("認証付き接続の確認に必要なYMM4 APIトークンがありません")
    # A positive response alone does not prove the endpoint enforces auth.
    invalid_token = ("0" if token[0] != "0" else "1") + token[1:]
    invalid_headers = {**headers, "X-Ymm4-Token": invalid_token}
    # Request an uncompressed response so HTTPX never expands a compressed
    # payload before the streaming size limit can be enforced.
    status_headers = {**headers, "Accept-Encoding": "identity"}
    url = base + "/status"
    # Both probes stay on loopback; ambient proxies must never see the token.
    # Stream before buffering so an unexpected local responder cannot force
    # an unbounded allocation. HTTPX's read timeout is per operation; also
    # enforce a total response deadline against indefinitely trickled data.
    deadline = time.monotonic() + 10.0
    with httpx.Client(timeout=5.0, trust_env=False) as client:
        # Inspect only the headers of the negative probe, without buffering
        # an untrusted response body.
        with client.stream("GET", url, headers=invalid_headers) as rejected:
            if rejected.status_code != 401:
                raise ValueError(
                    "YMM4の認証が有効であることを確認できません"
                    "（不正なトークンへの応答がHTTP 401ではありません）"
                )
        with client.stream("GET", url, headers=status_headers) as response:
            response.raise_for_status()
            # Do not allow gzip/deflate to inflate an arbitrarily large
            # response inside HTTPX before our byte-count check runs.
            content_encoding = response.headers.get("content-encoding")
            if isinstance(content_encoding, str) and content_encoding.strip().lower() not in ("", "identity"):
                raise ValueError("YMM4のステータス応答は圧縮できません")
            # If the responder declares a media type, require JSON. An HTML
            # error page or an incompatible charset must never pass as status.
            # A missing Content-Type remains supported for older YMM4 hosts.
            content_type = response.headers.get("content-type")
            if content_type is not None:
                parts = content_type.split(";")
                if parts[0].strip().lower() != "application/json":
                    raise ValueError("YMM4のステータスContent-TypeがJSONではありません")
                if len(parts) > 2:
                    raise ValueError("YMM4のステータスContent-Typeが不正です")
                if len(parts) == 2:
                    key, separator, value = parts[1].partition("=")
                    if (not separator or key.strip().lower() != "charset"
                            or value.strip().strip('"').lower() not in ("utf-8", "utf8")):
                        raise ValueError("YMM4のステータスContent-Typeの文字コードが不正です")
            # Content-Length is optional for chunked responses. When present,
            # reject oversized, malformed, or inconsistent framing.
            length_header = response.headers.get("content-length")
            transfer_encoding = response.headers.get("transfer-encoding")
            if transfer_encoding is not None:
                # RFC framing forbids using Content-Length alongside
                # Transfer-Encoding. Reject ambiguous/multi-coded responses.
                if (not isinstance(transfer_encoding, str)
                        or transfer_encoding.strip().lower() != "chunked"):
                    raise ValueError("YMM4のステータスTransfer-Encodingが不正です")
                if length_header is not None:
                    raise ValueError("YMM4のステータスHTTPフレーミングが矛盾しています")
            declared_length = None
            if length_header is not None:
                if (not isinstance(length_header, str) or not length_header.isascii()
                        or not length_header.isdigit() or len(length_header) > 5):
                    raise ValueError("YMM4のステータスContent-Lengthが不正です")
                declared_length = int(length_header)
                if declared_length > MAX_STATUS_JSON_BYTES:
                    raise ValueError("YMM4のステータスJSONが大きすぎます")
            body = bytearray()
            # Do not request fixed-size chunks: HTTPX may buffer a trickle
            # of tiny network reads until 8192 bytes arrive, bypassing the
            # total deadline despite each read staying below its timeout.
            for chunk in response.iter_bytes():
                if time.monotonic() > deadline:
                    raise ValueError("YMM4のステータス応答がタイムアウトしました")
                if len(body) + len(chunk) > MAX_STATUS_JSON_BYTES:
                    raise ValueError("YMM4のステータスJSONが大きすぎます")
                body.extend(chunk)
            if declared_length is not None and len(body) != declared_length:
                raise ValueError("YMM4のステータスContent-Lengthが応答サイズと一致しません")
    # json.loads(bytes) detects UTF-8/16/32. Decode with the same rules
    # before counting nesting; scanning raw UTF-16 bytes misses escapes.
    raw = bytes(body)
    decoded = raw.decode(json.detect_encoding(raw))
    if _status_json_too_deep(decoded):
        raise ValueError("YMM4のステータスJSONが深すぎます")
    try:
        status = json.loads(
            decoded,
            object_pairs_hook=_unique_json_object,
            parse_constant=_reject_nonfinite_json_constant,
            parse_float=_finite_json_float,
            parse_int=_bounded_json_int,
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
