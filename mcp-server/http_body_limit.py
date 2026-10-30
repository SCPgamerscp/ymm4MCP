"""Bound authenticated Streamable HTTP request bodies before MCP dispatch.

ASGI servers need not enforce Content-Length and can deliver chunked bodies without
one. This middleware caps received bytes before the MCP JSON parser sees any input,
while preserving the receive channel for disconnect notifications.
"""
from starlette.responses import JSONResponse

MAX_REQUEST_BYTES = 16 * 1024 * 1024
# A malicious client can send an unlimited number of tiny/empty ASGI frames
# without ever exceeding the byte limit. Bound CPU work as well as memory.
MAX_REQUEST_EVENTS = 65_536


class RequestBodyLimit:
    def __init__(self, app, max_bytes=MAX_REQUEST_BYTES):
        if isinstance(max_bytes, bool) or not isinstance(max_bytes, int) or max_bytes <= 0:
            raise ValueError("max_bytes must be a positive integer")
        self.app = app
        self.max_bytes = max_bytes

    async def __call__(self, scope, receive, send):
        if scope["type"] != "http" or scope.get("method", "").upper() != "POST":
            return await self.app(scope, receive, send)

        headers = scope.get("headers", [])
        lengths = [value for name, value in headers if name.lower() == b"content-length"]
        # A proxy and the ASGI server may disagree about framing when both
        # Content-Length and Transfer-Encoding are supplied.
        has_transfer_encoding = any(name.lower() == b"transfer-encoding"
                                    for name, _ in headers)
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
            return await receive()

        return await self.app(scope, replay_receive, send)
