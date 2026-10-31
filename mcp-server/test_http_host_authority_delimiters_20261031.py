"""Regression tests for ambiguous HTTP Host authorities."""
import asyncio
import unittest

from http_body_limit import RequestBodyLimit


class HostAuthorityTests(unittest.TestCase):
    def test_rejects_ambiguous_authorities_before_dispatch(self):
        async def run(host):
            called = []
            messages = []

            async def app(scope, receive, send):
                called.append(True)

            async def receive():
                return {"type": "http.request", "body": b"{}", "more_body": False}

            async def send(message):
                messages.append(message)

            scope = {
                "type": "http",
                "method": "POST",
                "http_version": "1.1",
                "headers": [(b"host", host)],
            }
            await RequestBodyLimit(app)(scope, receive, send)
            return called, next(m["status"] for m in messages if m["type"] == "http.response.start")

        for host in (b"example.com ", b" example.com", b"example.com\t", b"foo bar",
                     b"example.com/path", b"example.com?query", b"example.com#fragment",
                     b"user@example.com", b"example.com\x7f",
                     b"example.com\x80", b"example.com\xff", b"example.com:abc", b"example.com:",
                     b"example.com:80:90", b"[::1", b"[::1]suffix",
                     b"[::1]:bad", b"[::1]:", b"foo]bar", b"[foo]"):
            with self.subTest(host=host):
                called, status = asyncio.run(run(host))
                self.assertFalse(called)
                self.assertEqual(status, 400)

    def test_valid_host_is_forwarded(self):
        async def run(host):
            called = []
            async def app(scope, receive, send):
                called.append(await receive())
            async def receive():
                return {"type": "http.request", "body": b"{}", "more_body": False}
            async def send(message):
                pass
            await RequestBodyLimit(app)({
                "type": "http", "method": "POST", "http_version": "1.1",
                "headers": [(b"host", host)]
            }, receive, send)
            return called

        for host in (b"example.com", b"example.com:8080", b"[::1]:8080"):
            with self.subTest(host=host):
                self.assertEqual(asyncio.run(run(host))[0]["body"], b"{}")


if __name__ == "__main__":
    unittest.main()
