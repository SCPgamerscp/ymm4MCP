"""Regression tests for empty DNS labels in Host authorities."""
import asyncio
import importlib.util
from pathlib import Path
import unittest

MODULE_PATH = Path(__file__).with_name("http_body_limit.py")
spec = importlib.util.spec_from_file_location("http_body_limit", MODULE_PATH)
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)


async def request(host):
    called = []
    messages = [{"type": "http.request", "body": b"{}", "more_body": False}]

    async def app(scope, receive, send):
        called.append(True)
        await receive()
        await send({"type": "http.response.start", "status": 200, "headers": []})
        await send({"type": "http.response.body", "body": b"ok"})

    async def receive():
        return messages.pop(0)

    async def send(message):
        if message["type"] == "http.response.start":
            called.append(message["status"])

    scope = {"type": "http", "method": "POST", "http_version": "1.1",
             "headers": [(b"host", host)]}
    await module.RequestBodyLimit(app)(scope, receive, send)
    return called


class HostEmptyLabelTests(unittest.TestCase):
    def test_reject_empty_labels(self):
        for host in (b".", b".example.com", b"example..com", b"example.com..",
                     b"localhost..:8000", b"example..com:443"):
            with self.subTest(host=host):
                self.assertEqual(asyncio.run(request(host)), [400])

    def test_accept_normal_and_trailing_dot(self):
        for host in (b"localhost", b"example.com", b"example.com.",
                     b"example.com:443", b"[::1]:8000"):
            with self.subTest(host=host):
                self.assertEqual(asyncio.run(request(host)), [True, 200])


if __name__ == "__main__":
    unittest.main()
