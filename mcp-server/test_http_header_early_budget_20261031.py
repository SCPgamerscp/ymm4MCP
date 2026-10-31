"""Regression tests for early header byte-budget rejection (Issue #13)."""
import asyncio
import importlib.util
import pathlib
import unittest

MODULE_PATH = pathlib.Path(__file__).with_name("http_body_limit.py")
spec = importlib.util.spec_from_file_location("http_body_limit_early_header_budget", MODULE_PATH)
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)


class HeaderBudgetTests(unittest.TestCase):
    def _run(self, headers):
        called = []
        sent = []

        async def app(scope, receive, send):
            called.append(True)
            await receive()

        async def receive():
            return {"type": "http.request", "body": b"", "more_body": False}

        async def send(message):
            sent.append(message)

        scope = {"type": "http", "method": "POST", "headers": headers}
        asyncio.run(module.RequestBodyLimit(app)(scope, receive, send))
        return called, next((event["status"] for event in sent if event["type"] == "http.response.start"), None)

    def test_oversized_header_value_returns_431_before_dispatch(self):
        called, status = self._run([(b"x-test", b"z" * (module.MAX_HEADER_BYTES + 1))])
        self.assertFalse(called)
        self.assertEqual(status, 431)

    def test_aggregate_header_budget_returns_431(self):
        called, status = self._run([(b"x-one", b"a" * 40000), (b"x-two", b"b" * 30000)])
        self.assertFalse(called)
        self.assertEqual(status, 431)

    def test_exact_header_budget_is_allowed(self):
        name = b"x-test"
        called, status = self._run([(name, b"a" * (module.MAX_HEADER_BYTES - len(name)))])
        self.assertTrue(called)
        self.assertIsNone(status)


if __name__ == "__main__":
    unittest.main()
