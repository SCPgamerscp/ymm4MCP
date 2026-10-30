"""Regression tests for ambiguous Host authorities on authenticated MCP POST."""
import asyncio
import importlib.util
import pathlib
import unittest

path = pathlib.Path(__file__).with_name("http_body_limit.py")
spec = importlib.util.spec_from_file_location("http_body_limit_host_authority", path)
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)


class HostAuthorityTests(unittest.TestCase):
    def invoke(self, headers):
        dispatched, sent = [], []

        async def app(scope, receive, send):
            dispatched.append(await receive())

        async def receive():
            return {"type": "http.request", "body": b"{}", "more_body": False}

        async def send(message):
            sent.append(message)

        asyncio.run(module.RequestBodyLimit(app)(
            {"type": "http", "method": "POST", "headers": headers},
            receive, send))
        return dispatched, [m["status"] for m in sent if m["type"] == "http.response.start"]

    def test_duplicate_host_rejected_case_insensitively(self):
        dispatched, statuses = self.invoke([(b"Host", b"example.test"), (b"host", b"other.test")])
        self.assertFalse(dispatched)
        self.assertEqual(statuses, [400])

    def test_comma_joined_host_rejected(self):
        dispatched, statuses = self.invoke([(b"host", b"example.test,other.test")])
        self.assertFalse(dispatched)
        self.assertEqual(statuses, [400])

    def test_blank_host_rejected(self):
        dispatched, statuses = self.invoke([(b"host", b"  ")])
        self.assertFalse(dispatched)
        self.assertEqual(statuses, [400])

    def test_single_host_accepted(self):
        dispatched, statuses = self.invoke([(b"host", b"example.test:8080")])
        self.assertEqual(len(dispatched), 1)
        self.assertEqual(statuses, [])

    def test_missing_host_preserves_legacy_scope(self):
        dispatched, statuses = self.invoke([])
        self.assertEqual(len(dispatched), 1)
        self.assertEqual(statuses, [])


if __name__ == "__main__":
    unittest.main()
