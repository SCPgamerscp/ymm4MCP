"""HTTP method case-sensitivity regressions for Issue #13."""
import asyncio
import importlib.util
import pathlib
import unittest

path = pathlib.Path(__file__).with_name("http_body_limit.py")
spec = importlib.util.spec_from_file_location("http_body_limit_method_case", path)
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)


class MethodCaseTests(unittest.TestCase):
    def _invoke(self, method):
        called, sent = [], []

        async def app(scope, receive, send):
            called.append(True)

        async def receive():
            return {"type": "http.request", "body": b"", "more_body": False}

        async def send(event):
            sent.append(event)

        asyncio.run(module.RequestBodyLimit(app)(
            {"type": "http", "method": method, "headers": []}, receive, send))
        statuses = [event["status"] for event in sent if event["type"] == "http.response.start"]
        return called, statuses

    def test_lowercase_post_is_rejected(self):
        called, statuses = self._invoke("post")
        self.assertFalse(called)
        self.assertEqual(statuses, [400])

    def test_mixed_case_post_is_rejected(self):
        called, statuses = self._invoke("PoSt")
        self.assertFalse(called)
        self.assertEqual(statuses, [400])

    def test_canonical_post_is_dispatched(self):
        called, statuses = self._invoke("POST")
        self.assertTrue(called)
        self.assertEqual(statuses, [])

    def test_unrelated_method_is_preserved(self):
        called, statuses = self._invoke("GET")
        self.assertTrue(called)
        self.assertEqual(statuses, [])


if __name__ == "__main__":
    unittest.main()
