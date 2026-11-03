"""Regression: non-POST requests must not bypass Host authority validation."""
import asyncio
import unittest

from http_body_limit import RequestBodyLimit


class NonPostHostValidationTests(unittest.TestCase):
    def test_invalid_host_is_rejected_for_get_and_delete(self):
        async def exercise(method, host):
            called = []
            statuses = []

            async def downstream(scope, receive, send):
                called.append(True)
                await send({"type": "http.response.start", "status": 204, "headers": []})
                await send({"type": "http.response.body", "body": b""})

            async def receive():
                return {"type": "http.request", "body": b"", "more_body": False}

            async def send(message):
                if message["type"] == "http.response.start":
                    statuses.append(message["status"])

            scope = {"type": "http", "method": method, "path": "/", "headers": [(b"host", host)]}
            await RequestBodyLimit(downstream)(scope, receive, send)
            return statuses, called

        for method in ("GET", "DELETE"):
            for host in (b"bad..example", b"example.com:abc", b"[::1", b"example.com,evil.test"):
                with self.subTest(method=method, host=host):
                    status, called = asyncio.run(exercise(method, host))
                    self.assertEqual(status, [400])
                    self.assertEqual(called, [])

    def test_valid_get_reaches_downstream(self):
        async def exercise():
            statuses = []
            async def downstream(scope, receive, send):
                await send({"type": "http.response.start", "status": 204, "headers": []})
                await send({"type": "http.response.body", "body": b""})
            async def receive():
                return {"type": "http.request", "body": b"", "more_body": False}
            async def send(message):
                if message["type"] == "http.response.start":
                    statuses.append(message["status"])
            await RequestBodyLimit(downstream)(
                {"type": "http", "method": "GET", "path": "/", "headers": [(b"host", b"localhost:8000")]},
                receive, send)
            return statuses
        self.assertEqual(asyncio.run(exercise()), [204])


if __name__ == "__main__":
    unittest.main()
