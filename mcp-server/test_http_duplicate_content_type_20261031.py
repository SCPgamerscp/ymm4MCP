"""Regression tests for duplicate Content-Type header rejection."""
import asyncio
import importlib.util
from pathlib import Path
import unittest

spec = importlib.util.spec_from_file_location("http_body_limit", Path(__file__).with_name("http_body_limit.py"))
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)


async def dispatch(headers):
    called = []
    messages = [{"type": "http.request", "body": b"{}", "more_body": False}]

    async def app(scope, receive, send):
        called.append("app")
        await receive()
        await send({"type": "http.response.start", "status": 200, "headers": []})
        await send({"type": "http.response.body", "body": b"ok"})

    async def receive():
        return messages.pop(0)

    async def send(message):
        if message["type"] == "http.response.start":
            called.append(message["status"])

    scope = {"type": "http", "method": "POST", "http_version": "1.1", "headers": headers}
    await module.RequestBodyLimit(app)(scope, receive, send)
    return called


class DuplicateContentTypeTests(unittest.TestCase):
    def test_duplicate_same_case(self):
        headers = [(b"content-type", b"application/json"), (b"content-type", b"text/plain")]
        self.assertEqual(asyncio.run(dispatch(headers)), [400])

    def test_duplicate_mixed_case(self):
        headers = [(b"Content-Type", b"application/json"), (b"content-type", b"application/json")]
        self.assertEqual(asyncio.run(dispatch(headers)), [400])

    def test_comma_joined_content_type_rejected(self):
        self.assertEqual(asyncio.run(dispatch([
            (b"content-type", b"application/json, text/plain")
        ])), [400])

    def test_empty_content_type_rejected(self):
        self.assertEqual(asyncio.run(dispatch([(b"content-type", b"  \t")])), [400])

    def test_quoted_parameter_comma_allowed(self):
        self.assertEqual(asyncio.run(dispatch([
            (b"content-type", b'application/json; profile="a,b"')
        ])), ["app", 200])

    def test_escaped_quote_and_quoted_comma_allowed(self):
        self.assertEqual(asyncio.run(dispatch([
            (b"content-type", b'application/json; profile="a\\\"b,c"')
        ])), ["app", 200])

    def test_unquoted_comma_after_quoted_parameter_rejected(self):
        self.assertEqual(asyncio.run(dispatch([
            (b"content-type", b'application/json; profile="a,b", text/plain')
        ])), [400])

    def test_unclosed_quote_rejected(self):
        self.assertEqual(asyncio.run(dispatch([
            (b"content-type", b'application/json; profile="unterminated')
        ])), [400])

    def test_invalid_media_type_syntax(self):
        for value in (b"application", b"/json", b"application/", b"application//json",
                      b"application /json", b"application/json extra", b"; charset=utf-8",
                      b"application/@json"):
            with self.subTest(value=value):
                self.assertEqual(asyncio.run(dispatch([(b"content-type", value)])), [400])

    def test_valid_media_type_tokens(self):
        for value in (b"application/json", b"application/vnd.example+json",
                      b"text/plain; charset=utf-8", b"application/*"):
            with self.subTest(value=value):
                self.assertEqual(asyncio.run(dispatch([(b"content-type", value)])), ["app", 200])

    def test_single_content_type_allowed(self):
        self.assertEqual(asyncio.run(dispatch([(b"content-type", b"application/json")])), ["app", 200])

    def test_absent_content_type_unchanged(self):
        self.assertEqual(asyncio.run(dispatch([])), ["app", 200])


if __name__ == "__main__":
    unittest.main()
