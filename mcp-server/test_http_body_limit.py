"""ASGI regression tests for bounded authenticated MCP request bodies."""
import unittest

import httpx
from starlette.applications import Starlette
from starlette.requests import Request
from starlette.responses import JSONResponse
from starlette.routing import Route

from http_body_limit import RequestBodyLimit


async def echo(request: Request):
    data = await request.body()
    return JSONResponse({"length": len(data), "payload": data.decode("ascii")})


async def async_next(iterator):
    return next(iterator)


async def async_return(value):
    return value


class RequestBodyLimitTests(unittest.IsolatedAsyncioTestCase):
    def app(self, limit=12):
        return RequestBodyLimit(Starlette(routes=[Route("/mcp", echo, methods=["POST"])]),
                                max_bytes=limit)

    async def test_accepts_body_at_exact_limit_and_preserves_payload(self):
        async with httpx.AsyncClient(transport=httpx.ASGITransport(app=self.app()),
                                     base_url="https://localhost") as client:
            response = await client.post("/mcp", content=b"hello world!")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json(), {"length": 12, "payload": "hello world!"})

    async def test_rejects_oversized_declared_body_before_receiving(self):
        messages = []
        async def no_receive():
            raise AssertionError("oversized Content-Length must not read body")
        async def send(message):
            messages.append(message)
        await self.app()({"type": "http", "method": "POST", "headers": [
            (b"content-length", b"13")]}, no_receive, send)
        self.assertEqual(messages[0]["status"], 413)

    async def test_rejects_oversized_chunked_body_without_length(self):
        async def stream():
            yield b"hello "
            yield b"world!!"
        async with httpx.AsyncClient(transport=httpx.ASGITransport(app=self.app()),
                                     base_url="https://localhost") as client:
            response = await client.post("/mcp", content=stream())
        self.assertEqual(response.status_code, 413)

    async def test_rejects_ambiguous_or_malformed_content_length(self):
        for values in ([b"1", b"1"], [b"not-a-number"], [b"-1"], [b" 1"], [b"1,1"], [b"9" * 10000]):
            with self.subTest(values=values):
                messages = []
                async def no_receive():
                    raise AssertionError("bad headers must not read body")
                async def send(message):
                    messages.append(message)
                await self.app()({"type": "http", "method": "POST", "headers": [
                    (b"content-length", value) for value in values]}, no_receive, send)
                self.assertEqual(messages[0]["status"], 400)

    async def test_rejects_mismatched_declared_body_length(self):
        for declared in (2, 4):
            with self.subTest(declared=declared):
                messages = []
                async def receive():
                    return {"type": "http.request", "body": b"abc", "more_body": False}
                async def send(message):
                    messages.append(message)
                await self.app()({"type": "http", "method": "POST", "headers": [
                    (b"content-length", str(declared).encode())]}, receive, send)
                self.assertEqual(messages[0]["status"], 400)

    async def test_disconnect_does_not_dispatch_partial_request(self):
        dispatched = False
        async def inner(scope, receive, send):
            nonlocal dispatched
            dispatched = True
        app = RequestBodyLimit(inner, max_bytes=12)
        calls = iter([{"type": "http.request", "body": b"partial", "more_body": True},
                      {"type": "http.disconnect"}])
        await app({"type": "http", "method": "POST", "headers": []},
                  lambda: async_next(calls), lambda _: None)
        self.assertFalse(dispatched)

    async def test_other_methods_do_not_consume_receive(self):
        async def inner(scope, receive, send):
            return await receive()
        app = RequestBodyLimit(inner, max_bytes=12)
        expected = {"type": "http.disconnect"}
        result = await app({"type": "http", "method": "GET", "headers": []},
                           lambda: async_return(expected), lambda _: None)
        self.assertEqual(result, expected)


if __name__ == "__main__":
    unittest.main()
