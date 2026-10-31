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

    def test_reject_overlong_dns_names(self):
        for host in (b"a" * 64 + b".example", b"a" * 254,
                     b".".join([b"a" * 63] * 4)):
            with self.subTest(host=host):
                self.assertEqual(asyncio.run(request(host)), [400])

    def test_accept_dns_length_boundaries(self):
        for host in (b"a" * 63 + b".example",
                     b".".join([b"a" * 63] * 3 + [b"a" * 61]),
                     b".".join([b"a" * 63] * 3 + [b"a" * 61]) + b"."):
            with self.subTest(host=host):
                self.assertEqual(asyncio.run(request(host)), [True, 200])

    def test_reject_invalid_dns_label_syntax(self):
        for host in (b"-example.com", b"example-.com", b"foo.-bar",
                     b"foo.bar-", b"foo_bar.com", b"foo!bar.com",
                     b"foo+bar.com", b"foo=bar.com", b"foo:bar:443"):
            with self.subTest(host=host):
                self.assertEqual(asyncio.run(request(host)), [400])

    def test_accept_interior_hyphens_and_digits(self):
        for host in (b"a-b.example", b"123.example", b"a1-b2.example:443",
                     b"xn--bcher-kva.example", b"localhost."):
            with self.subTest(host=host):
                self.assertEqual(asyncio.run(request(host)), [True, 200])

    def test_accept_normal_and_trailing_dot(self):
        for host in (b"localhost", b"example.com", b"example.com.",
                     b"example.com:443", b"[::1]:8000"):
            with self.subTest(host=host):
                self.assertEqual(asyncio.run(request(host)), [True, 200])


if __name__ == "__main__":
    unittest.main()
