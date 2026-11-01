"""Regression: non-POST requests must not bypass the HTTP header budget."""
import asyncio
import importlib.util
from pathlib import Path
import unittest

MODULE = Path(__file__).with_name("http_body_limit.py")
spec = importlib.util.spec_from_file_location("http_body_limit", MODULE)
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)


class HeaderBudgetForAllMethods(unittest.TestCase):
    def invoke(self, method, headers):
        called = []
        sent = []

        async def app(scope, receive, send):
            called.append(True)
            await send({"type": "http.response.start", "status": 204, "headers": []})
            await send({"type": "http.response.body", "body": b""})

        async def receive():
            return {"type": "http.request", "body": b"", "more_body": False}

        async def send(message):
            sent.append(message)

        scope = {"type": "http", "method": method, "headers": headers, "http_version": "1.1"}
        asyncio.run(module.RequestBodyLimit(app)(scope, receive, send))
        return sent[0]["status"], bool(called)

    def test_get_rejects_oversized_header(self):
        status, called = self.invoke("GET", [(b"x-test", b"a" * (module.MAX_HEADER_BYTES + 1))])
        self.assertEqual((status, called), (431, False))

    def test_delete_rejects_invalid_field_name(self):
        status, called = self.invoke("DELETE", [(b"bad name", b"value")])
        self.assertEqual((status, called), (400, False))

    def test_get_rejects_excessive_header_count(self):
        status, called = self.invoke("GET", [(b"x", b"y")] * (module.MAX_HEADER_FIELDS + 1))
        self.assertEqual((status, called), (400, False))

    def test_valid_get_reaches_downstream(self):
        status, called = self.invoke("GET", [(b"host", b"localhost")])
        self.assertEqual((status, called), (204, True))


if __name__ == "__main__":
    unittest.main()
