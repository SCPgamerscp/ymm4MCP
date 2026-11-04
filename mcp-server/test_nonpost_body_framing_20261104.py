"""Regression tests for non-POST request framing at the MCP body limiter."""
import asyncio
import unittest

from http_body_limit import RequestBodyLimit


class NonPostBodyFramingTests(unittest.TestCase):
    def test_get_delete_reject_framed_bodies_before_dispatch(self):
        for method in ("GET", "DELETE"):
            for framing in ((b"content-length", b"5"), (b"transfer-encoding", b"chunked")):
                with self.subTest(method=method, framing=framing):
                    called = []
                    status = []
                    async def downstream(scope, receive, send):
                        called.append(True)
                    async def receive():
                        return {"type": "http.request", "body": b"", "more_body": False}
                    async def send(message):
                        if message["type"] == "http.response.start":
                            status.append(message["status"])
                    scope = {"type": "http", "method": method, "scheme": "https",
                             "http_version": "1.1",
                             "headers": [(b"host", b"localhost"), framing]}
                    asyncio.run(RequestBodyLimit(downstream)(scope, receive, send))
                    self.assertEqual(status, [400])
                    self.assertFalse(called)

    def test_bodyless_get_delete_still_dispatch(self):
        for method in ("GET", "DELETE"):
            with self.subTest(method=method):
                called = []
                async def downstream(scope, receive, send):
                    called.append(True)
                async def receive():
                    raise AssertionError("Bodyless method should not read request stream")
                async def send(message):
                    pass
                scope = {"type": "http", "method": method, "scheme": "https",
                         "headers": [(b"host", b"localhost")]}
                asyncio.run(RequestBodyLimit(downstream)(scope, receive, send))
                self.assertTrue(called)


if __name__ == "__main__":
    unittest.main()
