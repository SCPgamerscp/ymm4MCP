"""Bound authenticated Streamable HTTP request bodies before MCP dispatch.

ASGI servers need not enforce Content-Length and can deliver chunked bodies without
one. This middleware caps received bytes before the MCP JSON parser sees any input,
while preserving the receive channel for disconnect notifications.
"""
from starlette.responses import JSONResponse

MAX_REQUEST_BYTES = 16 * 1024 * 1024


class RequestBodyLimit:
    def __init__(self, app, max_bytes=MAX_REQUEST_BYTES):
        if isinstance(max_bytes, bool) or not isinstance(max_bytes, int) or max_bytes <= 0:
            raise ValueError("max_bytes must be a positive integer")
        self.app = app
        self.max_bytes = max_bytes

    async def __call__(self, scope, receive, send):
        if scope["type"] != "http" or scope.get("method", "").upper() != "POST":
            return await self.app(scope, receive, send)

        lengths = [value for name, value in scope.get("headers", [])
                   if name.lower() == b"content-length"]
        if len(lengths) > 1 or (lengths and
                (len(lengths[0]) > 20 or not lengths[0].isascii() or not lengths[0].isdigit())):
            return await JSONResponse({"error": "Invalid Content-Length"}, status_code=400)(
                scope, receive, send)
        declared = int(lengths[0]) if lengths else None
        if declared is not None and declared > self.max_bytes:
            return await JSONResponse({"error": "MCP request body too large"}, status_code=413)(
                scope, receive, send)

        body = bytearray()
        while True:
            message = await receive()
            if message["type"] == "http.disconnect":
                return
            if message["type"] != "http.request":
                return await JSONResponse({"error": "Invalid request stream"}, status_code=400)(
                    scope, receive, send)
            chunk = message.get("body", b"")
            if len(chunk) > self.max_bytes - len(body):
                return await JSONResponse({"error": "MCP request body too large"}, status_code=413)(
                    scope, receive, send)
            body.extend(chunk)
            if not message.get("more_body", False):
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
