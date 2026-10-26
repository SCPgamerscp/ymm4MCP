"""Reject ambiguous Host headers before MCP transport security validation (#22)."""
import os
import unittest
from unittest.mock import AsyncMock, patch

import httpx

from http_transport import create_http_app
import server


TOKEN = "test-only-not-for-production-0123456789"


class HostHeaderAmbiguityTests(unittest.IsolatedAsyncioTestCase):
    async def test_duplicate_host_headers_are_rejected_before_mcp_dispatch(self):
        with patch.dict(os.environ, {"YMM4_MCP_BEARER_TOKEN": TOKEN}, clear=True):
            app = create_http_app(server.app, AsyncMock())
        async with httpx.AsyncClient(
            transport=httpx.ASGITransport(app=app), base_url="https://localhost"
        ) as client:
            for hosts in (
                [("Host", "localhost"), ("Host", "evil.example")],
                [("Host", "localhost"), ("host", "localhost")],
            ):
                with self.subTest(hosts=hosts):
                    response = await client.post(
                        "/mcp",
                        headers=hosts + [("Authorization", f"Bearer {TOKEN}")],
                    )
                    self.assertEqual(response.status_code, 400)
                    self.assertEqual(
                        response.json()["error"], "Exactly one Host header is required"
                    )

    async def test_single_host_still_requires_bearer_authentication(self):
        with patch.dict(os.environ, {"YMM4_MCP_BEARER_TOKEN": TOKEN}, clear=True):
            app = create_http_app(server.app, AsyncMock())
        async with httpx.AsyncClient(
            transport=httpx.ASGITransport(app=app), base_url="https://localhost"
        ) as client:
            response = await client.post(
                "/mcp", headers=[("Host", "localhost"), ("Authorization", "Bearer wrong")]
            )
            self.assertEqual(response.status_code, 401)


if __name__ == "__main__":
    unittest.main()
