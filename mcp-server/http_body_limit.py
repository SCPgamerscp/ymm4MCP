"""Bound authenticated Streamable HTTP request bodies before MCP dispatch.

ASGI servers need not enforce Content-Length and can deliver chunked bodies without
one. This middleware caps received bytes before the MCP JSON parser sees any input,
while preserving the receive channel for disconnect notifications.
"""
from ipaddress import IPv6Address

from starlette.responses import JSONResponse

MAX_REQUEST_BYTES = 16 * 1024 * 1024
# A malicious client can send an unlimited number of tiny/empty ASGI frames
# without ever exceeding the byte limit. Bound CPU work as well as memory.
MAX_REQUEST_EVENTS = 65_536
# Reject pathological zero-progress streams earlier than the overall event cap.
MAX_EMPTY_REQUEST_EVENTS = 1_024
MAX_HEADER_FIELDS = 256
MAX_HEADER_BYTES = 64 * 1024
# RFC 9110 field names are ASCII tokens, not arbitrary header text.
_HEADER_TOKEN_BYTES = frozenset(b"!#$%&'*+-.^_`|~0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz")

# DNS host labels use ASCII letters, digits and interior hyphens.
_DNS_ALNUM_BYTES = frozenset(b"0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz")
_DNS_LABEL_BYTES = _DNS_ALNUM_BYTES | {45}


def _invalid_media_type(value):
    """Reject missing or invalid type/subtype tokens before optional parameters."""
    media_type = value.split(b";", 1)[0].strip()
    parts = media_type.split(b"/")
    # Content-Type describes a concrete representation, not an Accept-style
    # media range. A full-token wildcard is invalid here, while an\n    # embedded asterisk is permitted by the HTTP token grammar.
    return (len(parts) != 2 or
            any(not part or part == b"*" or
                any(char not in _HEADER_TOKEN_BYTES for char in part)
                for part in parts))


def _invalid_content_type_list(value):
    """Detect joined media types without rejecting quoted parameter commas."""
    quoted = False
    escaped = False
    for char in value:
        if escaped:
            escaped = False
        elif quoted and char == 92:  # quoted-pair escape
            escaped = True
        elif char == 34:  # double quote
            quoted = not quoted
        elif char == 44 and not quoted:  # comma outside a quoted-string
            return True
    return quoted or escaped


def _invalid_content_type_parameters(value):
    """Validate parameter names and values without splitting quoted semicolons."""
    if b";" not in value:
        return False
    quoted = False
    escaped = False
    start = value.index(b";") + 1
    parameters = []
    for index in range(start, len(value)):
        char = value[index]
        if escaped:
            escaped = False
        elif quoted and char == 92:
            escaped = True
        elif char == 34:
            quoted = not quoted
        elif char == 59 and not quoted:
            parameters.append(value[start:index])
            start = index + 1
    if quoted or escaped:
        return True
    parameters.append(value[start:])
    seen_names = set()
    for parameter in parameters:
        parameter = parameter.strip(b" \t")
        if b"=" not in parameter:
            return True
        name, raw_value = parameter.split(b"=", 1)
        name = name.strip(b" \t")
        raw_value = raw_value.strip(b" \t")
        if not name or any(char not in _HEADER_TOKEN_BYTES for char in name):
            return True
        # Duplicate parameter names are case-insensitive and ambiguous to consumers.
        canonical_name = name.lower()
        if canonical_name in seen_names:
            return True
        seen_names.add(canonical_name)
        if raw_value.startswith(b'"'):
            if len(raw_value) < 2 or not raw_value.endswith(b'"'):
                return True
            inner = raw_value[1:-1]
            escaped = False
            for char in inner:
                if escaped:
                    if (char < 32 and char != 9) or char == 127:
                        return True
                    escaped = False
                elif char == 92:
                    escaped = True
                elif char == 34 or (char < 32 and char != 9) or char == 127:
                    return True
            if escaped:
                return True
        elif not raw_value or any(char not in _HEADER_TOKEN_BYTES for char in raw_value):
            return True
    return False


class RequestBodyLimit:
    def __init__(self, app, max_bytes=MAX_REQUEST_BYTES):
        if isinstance(max_bytes, bool) or not isinstance(max_bytes, int) or max_bytes <= 0:
            raise ValueError("max_bytes must be a positive integer")
        self.app = app
        self.max_bytes = max_bytes

    async def __call__(self, scope, receive, send):
        if scope["type"] != "http":
            return await self.app(scope, receive, send)
        method = scope.get("method", "")
        if (not isinstance(method, str) or not method or not method.isascii() or
                any(char not in _HEADER_TOKEN_BYTES for char in method.encode("ascii"))):
            return await JSONResponse({"error": "Invalid request method"}, status_code=400)(
                scope, receive, send)
        # HTTP method tokens are case-sensitive (RFC 9110). Do not silently
        # reinterpret a different token such as "post" as POST.
        if method != "POST":
            if method.upper() == "POST":
                return await JSONResponse({"error": "Invalid request method"}, status_code=400)(
                    scope, receive, send)
            return await self.app(scope, receive, send)

        headers = scope.get("headers", [])
        if not isinstance(headers, (list, tuple)) or len(headers) > MAX_HEADER_FIELDS:
            return await JSONResponse({"error": "Invalid request headers"}, status_code=400)(
                scope, receive, send)
        normalized_headers = []
        header_bytes = 0
        for field in headers:
            if not isinstance(field, (list, tuple)) or len(field) != 2:
                return await JSONResponse({"error": "Invalid request headers"}, status_code=400)(
                    scope, receive, send)
            name, value = field
            if not isinstance(name, bytes) or not isinstance(value, bytes):
                return await JSONResponse({"error": "Invalid request headers"}, status_code=400)(
                    scope, receive, send)
            # Check sizes before scanning potentially attacker-sized header values.
            # The total budget includes names and values, not just values.
            if len(name) + len(value) > MAX_HEADER_BYTES - header_bytes:
                return await JSONResponse({"error": "Request headers too large"}, status_code=431)(
                    scope, receive, send)
            if (not name or any(char not in _HEADER_TOKEN_BYTES for char in name) or
                    any((char < 32 and char != 9) or char == 127 for char in value)):
                return await JSONResponse({"error": "Invalid request headers"}, status_code=400)(
                    scope, receive, send)
            header_bytes += len(name) + len(value)
            if header_bytes > MAX_HEADER_BYTES:
                return await JSONResponse({"error": "Request headers too large"}, status_code=431)(
                    scope, receive, send)
            normalized_headers.append((name.lower(), value))

        # Multiple Host fields or comma-joined authorities can be interpreted
        # differently by upstream proxies and the application router.
        hosts = [value for name, value in normalized_headers if name == b"host"]
        if len(hosts) > 1 or (hosts and (not hosts[0] or any(
                char <= 32 or char >= 127 or char in b",/@?#\\%" for char in hosts[0]))):
            return await JSONResponse({"error": "Invalid Host header"}, status_code=400)(
                scope, receive, send)

        # Reject malformed bracketed IPv6 authorities and nonnumeric ports.
        if hosts:
            authority = hosts[0]
            if authority.startswith(b"["):
                closing = authority.find(b"]")
                suffix = authority[closing + 1:] if closing >= 0 else b""
                valid_authority = (closing > 1 and (not suffix or (suffix.startswith(b":")
                                   and 1 <= len(suffix[1:]) <= 5 and suffix[1:].isdigit()
                                   and int(suffix[1:]) <= 65535)))
                if valid_authority:
                    try:
                        IPv6Address(authority[1:closing].decode("ascii"))
                    except (ValueError, UnicodeDecodeError):
                        valid_authority = False
            else:
                host, sep, port = authority.partition(b":")
                valid_authority = (bool(host) and len(host.rstrip(b".")) <= 253
                                   and all(1 <= len(label) <= 63 and label[0] in _DNS_ALNUM_BYTES
                                           and label[-1] in _DNS_ALNUM_BYTES
                                           and all(char in _DNS_LABEL_BYTES for char in label)
                                           for label in host.rstrip(b".").split(b"."))
                                   and host != b"." and not host.startswith(b".")
                                   and not host.endswith(b"..")
                                   and b"[" not in authority and b"]" not in authority
                                   and (not sep or (1 <= len(port) <= 5 and port.isdigit()
                                                    and int(port) <= 65535)))
            if not valid_authority:
                return await JSONResponse({"error": "Invalid Host header"}, status_code=400)(
                    scope, receive, send)

        # A duplicated Content-Type may be interpreted inconsistently by parsers.
        content_types = [value for name, value in normalized_headers if name == b"content-type"]
        if len(content_types) > 1:
            return await JSONResponse({"error": "Duplicate Content-Type"}, status_code=400)(
                scope, receive, send)

        # Content-Type is a singleton field, not a comma-separated list. Reject
        # joined values that could be split differently by downstream parsers.
        if content_types and (not content_types[0].strip(b" \t") or _invalid_media_type(content_types[0]) or _invalid_content_type_list(content_types[0]) or
                              _invalid_content_type_parameters(content_types[0])):
            return await JSONResponse({"error": "Invalid Content-Type"}, status_code=400)(
                scope, receive, send)

        lengths = [value for name, value in normalized_headers if name == b"content-length"]
        transfers = [value for name, value in normalized_headers if name == b"transfer-encoding"]
        # Only a single chunked encoding is supported. Reject ambiguous or
        # stacked transfer codings rather than trusting conflicting parsers.
        if len(transfers) > 1 or (transfers and transfers[0].strip(b" \t").lower() != b"chunked"):
            return await JSONResponse({"error": "Invalid Transfer-Encoding"}, status_code=400)(
                scope, receive, send)
        has_transfer_encoding = bool(transfers)
        # HTTP/2 and HTTP/3 do not use HTTP/1.1 chunked transfer coding.
        if has_transfer_encoding and scope.get("http_version", "1.1") != "1.1":
            return await JSONResponse({"error": "Invalid Transfer-Encoding"}, status_code=400)(
                scope, receive, send)
        if ((lengths and has_transfer_encoding) or len(lengths) > 1 or
                (lengths and (not isinstance(lengths[0], bytes) or
                              len(lengths[0]) > 20 or not lengths[0].isascii() or
                              not lengths[0].isdigit() or
                              (len(lengths[0]) > 1 and lengths[0].startswith(b"0"))))):
            return await JSONResponse({"error": "Invalid Content-Length"}, status_code=400)(
                scope, receive, send)
        declared = int(lengths[0]) if lengths else None
        if declared is not None and declared > self.max_bytes:
            return await JSONResponse({"error": "MCP request body too large"}, status_code=413)(
                scope, receive, send)

        body = bytearray()
        events = 0
        empty_events = 0
        while True:
            message = await receive()
            events += 1
            if not isinstance(message, dict):
                return await JSONResponse({"error": "Invalid request stream"}, status_code=400)(
                    scope, receive, send)
            if message.get("type") == "http.disconnect":
                return
            if message.get("type") != "http.request":
                return await JSONResponse({"error": "Invalid request stream"}, status_code=400)(
                    scope, receive, send)
            chunk = message.get("body", b"")
            more_body = message.get("more_body", False)
            if not isinstance(chunk, bytes) or not isinstance(more_body, bool):
                return await JSONResponse({"error": "Invalid request stream"}, status_code=400)(
                    scope, receive, send)
            if events > MAX_REQUEST_EVENTS:
                return await JSONResponse({"error": "Too many request chunks"}, status_code=413)(
                    scope, receive, send)
            if not chunk and more_body:
                empty_events += 1
                if empty_events > MAX_EMPTY_REQUEST_EVENTS:
                    return await JSONResponse({"error": "Too many empty request chunks"}, status_code=413)(
                        scope, receive, send)
            elif chunk:
                # The empty-frame threshold measures consecutive lack of progress.
                # A data-bearing frame resets it; the total event cap still applies.
                empty_events = 0
            if len(chunk) > self.max_bytes - len(body):
                return await JSONResponse({"error": "MCP request body too large"}, status_code=413)(
                    scope, receive, send)
            if declared is not None and len(chunk) > declared - len(body):
                return await JSONResponse({"error": "Content-Length mismatch"}, status_code=400)(
                    scope, receive, send)
            body.extend(chunk)
            if not more_body:
                break

        if declared is not None and declared != len(body):
            return await JSONResponse({"error": "Content-Length mismatch"}, status_code=400)(
                scope, receive, send)

        delivered = False

        async def replay_receive():
            nonlocal delivered
            if not delivered:
                delivered = True
                return {"type": "http.request", "body": bytes(body), "more_body": False}
            message = await receive()
            # Once more_body=False has been observed, a second request body
            # would bypass the size check. Only disconnect notifications are
            # valid after the replayed terminal request frame.
            if not isinstance(message, dict) or message.get("type") != "http.disconnect":
                return {"type": "http.disconnect"}
            return message

        return await self.app(scope, replay_receive, send)
