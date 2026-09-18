"""ASGI-level HTTPS MCP tests. No public listener is started."""
import asyncio
from contextlib import asynccontextmanager
import os
from types import SimpleNamespace
import unittest
from unittest.mock import AsyncMock, patch
import httpx
from http_transport import create_http_app, run_http
import server

TOKEN = "test-only-not-for-production-0123456789"


@asynccontextmanager
async def lifespan(app):
    incoming, outgoing = asyncio.Queue(), asyncio.Queue()
    task = asyncio.create_task(app({"type": "lifespan", "asgi": {"version": "3.0"}}, incoming.get, outgoing.put))
    try:
        await incoming.put({"type": "lifespan.startup"})
        message = await asyncio.wait_for(outgoing.get(), 5)
        if message["type"] != "lifespan.startup.complete":
            raise RuntimeError(message)
        yield
        await incoming.put({"type": "lifespan.shutdown"})
        message = await asyncio.wait_for(outgoing.get(), 5)
        if message["type"] != "lifespan.shutdown.complete":
            raise RuntimeError(message)
        await asyncio.wait_for(task, 5)
    finally:
        if not task.done():
            task.cancel()
        await asyncio.gather(task, return_exceptions=True)


class HttpTransportTests(unittest.IsolatedAsyncioTestCase):
    async def test_requires_credentials_and_tls(self):
        for env in ({}, {"YMM4_MCP_BEARER_TOKEN": "short"}, {"YMM4_MCP_BEARER_TOKEN": TOKEN, "YMM4_MCP_ALLOWED_HOSTS": "*"}):
            with self.subTest(env=env), patch.dict(os.environ, env, clear=True), self.assertRaises(ValueError):
                create_http_app(server.app, AsyncMock())
        with self.assertRaises(ValueError):
            await run_http(server.app, AsyncMock(), SimpleNamespace(tls_cert=None, tls_key=None))

    async def test_auth_tls_and_browser_rejections(self):
        with patch.dict(os.environ, {"YMM4_MCP_BEARER_TOKEN": TOKEN}, clear=True):
            app = create_http_app(server.app, AsyncMock())
        async with httpx.AsyncClient(transport=httpx.ASGITransport(app=app), base_url="https://localhost") as client:
            for path in ("/mcp", "/other"):
                response = await client.post(path)
                self.assertEqual(response.status_code, 401)
                self.assertNotIn("access-control-allow-origin", response.headers)
            response = await client.post("/mcp", headers={"Authorization": "Bearer wrong"})
            self.assertEqual(response.status_code, 401)
            response = await client.post("/mcp", headers={"Authorization": f"Bearer {TOKEN}", "Origin": "https://example.com"})
            self.assertEqual(response.status_code, 403)
            response = await client.post("http://localhost/mcp", headers={"Authorization": f"Bearer {TOKEN}"})
            self.assertEqual(response.status_code, 403)

    async def test_authenticated_mcp_and_host_validation(self):
        close = AsyncMock()
        with patch.dict(os.environ, {"YMM4_MCP_BEARER_TOKEN": TOKEN, "YMM4_ENABLE_ADVANCED": "0"}, clear=True):
            app = create_http_app(server.app, close)
            async with lifespan(app):
                async with httpx.AsyncClient(transport=httpx.ASGITransport(app=app), base_url="https://localhost", headers={"Authorization": f"Bearer {TOKEN}", "Accept": "application/json, text/event-stream"}) as client:
                    response = await client.post("/mcp", json={"jsonrpc": "2.0", "id": 1, "method": "initialize", "params": {"protocolVersion": "2025-06-18", "capabilities": {}, "clientInfo": {"name": "test", "version": "1"}}})
                    self.assertEqual(response.status_code, 200, response.text)
                    self.assertEqual(response.json()["result"]["serverInfo"]["name"], "ymm4-mcp")
                    for method, field, count in (("tools/list", "tools", 3), ("resources/list", "resources", 4), ("prompts/list", "prompts", 4)):
                        response = await client.post("/mcp", json={"jsonrpc": "2.0", "id": 2, "method": method})
                        self.assertEqual(response.status_code, 200, response.text)
                        self.assertEqual(len(response.json()["result"][field]), count)
                    response = await client.post("https://evil.example/mcp", json={"jsonrpc": "2.0", "id": 3, "method": "tools/list"})
                    self.assertIn(response.status_code, (400, 421))
            close.assert_awaited_once()


if __name__ == "__main__":
    unittest.main()
