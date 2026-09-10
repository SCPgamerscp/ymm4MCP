"""HTTP timeout/lifecycle regressions; no YMM4, sockets, or media writes needed.

Run: python -m unittest discover -s mcp-server -p 'test_http_timeouts.py' -v
"""

import asyncio
from contextlib import asynccontextmanager
import json
import unittest
from unittest.mock import AsyncMock, patch

import httpx

import server


class HttpTimeoutTests(unittest.IsolatedAsyncioTestCase):
    async def asyncSetUp(self):
        self.requests = []
        self.response = {"success": True, "image": "aW1hZ2U="}
        self.failure = None
        self.status_code = 200

        async def respond(request):
            self.requests.append(request)
            if self.failure is not None:
                raise self.failure
            return httpx.Response(self.status_code, json=self.response)

        self.client = httpx.AsyncClient(
            transport=httpx.MockTransport(respond), timeout=10.0
        )
        self.factory = patch.object(server.httpx, "AsyncClient", return_value=self.client)
        self.client_factory = self.factory.start()
        self.addCleanup(self.factory.stop)

    async def asyncTearDown(self):
        await server.close_http_client()
        await self.client.aclose()

    def assert_timeout(self, read):
        self.assertEqual(
            self.requests[-1].extensions["timeout"],
            {"connect": 10.0, "read": read, "write": 10.0, "pool": 10.0},
        )

    async def test_watch_and_record_use_effective_duration_plus_overhead(self):
        for action, default, minimum in (("watch", 5000, 1000), ("record", 3000, 500)):
            for duration in (None, 8000, 30000, -100, 0, 500, 1000, 60000):
                with self.subTest(action=action, duration=duration):
                    args = {"action": action, "frame": 42, "capture_interval_ms": 2000}
                    if duration is not None:
                        args["duration_ms"] = duration
                    result = await server.dispatch_preview(args)
                    self.assertFalse(result.isError)
                    sent_duration = default if duration is None else duration
                    self.assert_timeout(max(minimum, min(sent_duration, 30000)) / 1000 + 15)
                    request = self.requests[-1]
                    self.assertEqual(request.url.path, f"/api/preview/{action}")
                    payload = json.loads(request.content)
                    self.assertEqual(payload["duration_ms"], sent_duration)
                    if action == "watch":
                        self.assertEqual(payload["frame"], 42)
                        self.assertEqual(payload["capture_interval_ms"], 2000)
        self.client_factory.assert_called_once()

    async def test_invalid_duration_is_rejected_before_http_request(self):
        for action in ("watch", "record"):
            for duration in (None, "8000", 1.5, True, [], {}):
                with self.subTest(action=action, duration=duration):
                    result = await server.call_tool(
                        "ymm4_preview", {"action": action, "duration_ms": duration}
                    )
                    self.assertTrue(result.isError)
                    self.assertIn("duration_ms", result.content[0].text)
        self.assertEqual(self.requests, [])
        self.client_factory.assert_not_called()

    async def test_seek_has_rendering_overhead_and_returns_image(self):
        result = await server.dispatch_preview({"action": "seek_capture", "frame": 120})
        self.assert_timeout(15.0)
        self.assertEqual(json.loads(self.requests[-1].content), {"frame": 120})
        self.assertEqual(result.content[0].type, "image")
        self.assertEqual(result.content[0].data, "aW1hZ2U=")

    async def test_capture_and_position_keep_default_timeout(self):
        for action in ("capture", "position"):
            with self.subTest(action=action):
                result = await server.dispatch_preview({"action": action})
                self.assertFalse(result.isError)
                self.assert_timeout(10.0)

    async def test_long_posts_reuse_client_without_changing_other_timeouts(self):
        await server.ymm4_post_long("/preview/export-clip", {"startFrame": 0})
        self.assert_timeout(600.0)
        await server.ymm4_post_long("/preview/export-clip", {}, timeout=90.0)
        self.assert_timeout(90.0)
        await server.ymm4_post("/playback/stop")
        self.assert_timeout(10.0)
        self.assertEqual(json.loads(self.requests[-1].content), {})
        await server.ymm4_get("/status")
        self.assert_timeout(10.0)
        self.client_factory.assert_called_once()
        self.assertFalse(self.client.is_closed)

    async def test_parallel_requests_have_independent_timeouts(self):
        await asyncio.gather(
            server.dispatch_preview({"action": "watch", "duration_ms": 8000}),
            server.dispatch_preview({"action": "record", "duration_ms": 30000}),
            server.ymm4_get("/status"),
        )
        by_path = {r.url.path: r.extensions["timeout"]["read"] for r in self.requests}
        self.assertEqual(by_path, {
            "/api/preview/watch": 23.0,
            "/api/preview/record": 45.0,
            "/api/status": 10.0,
        })
        self.client_factory.assert_called_once()

    async def test_connection_errors_and_timeouts_are_distinct_and_not_retried(self):
        cases = (
            (httpx.ConnectError, "接続できません"),
            (httpx.ConnectTimeout, "接続できません"),
            (httpx.ReadTimeout, "処理待ちがタイムアウト"),
            (httpx.WriteTimeout, "処理待ちがタイムアウト"),
            (httpx.PoolTimeout, "処理待ちがタイムアウト"),
        )
        for error_type, expected in cases:
            with self.subTest(error=error_type.__name__):
                self.failure = error_type("")
                before = len(self.requests)
                result = await server.call_tool(
                    "ymm4_preview", {"action": "watch", "duration_ms": 8000}
                )
                self.assertTrue(result.isError)
                text = result.content[0].text
                self.assertIn(expected, text)
                self.assertEqual(len(self.requests), before + 1)
                if expected == "処理待ちがタイムアウト":
                    self.assertIn("自動再試行はしていません", text)
                    self.assertNotIn("接続できません", text)

    async def test_export_timeout_reaches_same_tool_error_handler(self):
        self.failure = httpx.ReadTimeout("")
        with patch.object(server.gemini_video, "is_available", return_value=(True, "")):
            result = await server.call_tool("ymm4_analyze_video", {"source": "preview"})
        self.assertTrue(result.isError)
        self.assertIn("処理待ちがタイムアウト", result.content[0].text)
        self.assertEqual(len(self.requests), 1)
        self.assert_timeout(600.0)

    async def test_http_status_errors_still_propagate(self):
        self.status_code = 500
        for method in (server.ymm4_get, server.ymm4_post, server.ymm4_post_long):
            with self.subTest(method=method.__name__):
                args = ("/status", {}) if method is server.ymm4_post_long else ("/status",)
                with self.assertRaises(httpx.HTTPStatusError) as caught:
                    await method(*args)
                self.assertEqual(caught.exception.response.status_code, 500)

    async def test_close_is_idempotent_and_next_session_gets_new_client(self):
        await server.ymm4_get("/status")
        await server.close_http_client()
        self.assertTrue(self.client.is_closed)
        self.assertIsNone(server._http_client)
        await server.close_http_client()
        # A fresh real client checks that the closed pool is never reused.
        replacement = type(self.client)(transport=httpx.MockTransport(
            lambda request: httpx.Response(200, json={"success": True})
        ))
        self.client_factory.return_value = replacement
        try:
            await server.ymm4_get("/status")
            self.assertIs(server._http_client, replacement)
            self.assertEqual(self.client_factory.call_count, 2)
        finally:
            await server.close_http_client()
            await replacement.aclose()

    async def test_main_closes_client_on_success_failure_and_cancellation(self):
        @asynccontextmanager
        async def streams():
            yield object(), object()

        for failure in (None, RuntimeError("MCP stopped"), asyncio.CancelledError()):
            with self.subTest(failure=type(failure).__name__):
                client = type(self.client)(transport=httpx.MockTransport(
                    lambda request: httpx.Response(200, json={"success": True})
                ))
                self.client_factory.return_value = client

                async def run(*args):
                    await server.ymm4_get("/status")
                    if failure is not None:
                        raise failure

                try:
                    with (
                        patch.object(server, "stdio_server", streams),
                        patch.object(server.app, "run", AsyncMock(side_effect=run)),
                    ):
                        if failure is None:
                            await server.main()
                        else:
                            with self.assertRaises(type(failure)):
                                await server.main()
                    self.assertTrue(client.is_closed)
                    self.assertIsNone(server._http_client)
                finally:
                    await server.close_http_client()
                    await client.aclose()


if __name__ == "__main__":
    unittest.main()
