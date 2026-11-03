"""Regression: non-POST requests must not bypass framing validation."""
import asyncio
import unittest

from http_body_limit import RequestBodyLimit


class NonPostFramingTests(unittest.TestCase):
    def test_invalid_framing_rejected_before_downstream(self):
        cases = [
            (b"content-length", b"1", b"transfer-encoding", b"chunked"),
            (b"content-length", b"1", b"content-length", b"2"),
            (b"content-length", b"01", None, None),
            (b"transfer-encoding", b"gzip", None, None),
        ]
        for method in ("GET", "DELETE"):
            for first_name, first_value, second_name, second_value in cases:
                with self.subTest(method=method, case=(first_name, first_value)):
                    headers = [(b"host", b"localhost"), (first_name, first_value)]
                    if second_name is not None:
                        headers.append((second_name, second_value))
                    self.assertEqual(self._run(method, headers), (400, False))

    def test_oversized_declared_body_rejected(self):
        self.assertEqual(self._run("GET", [
            (b"host", b"localhost"), (b"content-length", b"100")
        ], max_bytes=10), (413, False))

    def test_valid_get_still_reaches_downstream(self):
        self.assertEqual(self._run("GET", [
            (b"host", b"localhost"), (b"content-length", b"0")
        ]), (204, True))

    @staticmethod
    def _run(method, headers, max_bytes=1024):
        called = False
        status = None

        async def downstream(scope, receive, send):
            nonlocal called
            called = True
            await send({"type": "http.response.start", "status": 204, "headers": []})
            await send({"type": "http.response.body", "body": b""})

        async def receive():
            return {"type": "http.request", "body": b"", "more_body": False}

        async def send(message):
            nonlocal status
            if message["type"] == "http.response.start":
                status = message["status"]

        scope = {"type": "http", "method": method, "headers": headers,
                 "http_version": "1.1", "path": "/", "root_path": "",
                 "query_string": b"", "scheme": "http", "server": ("localhost", 80)}
        asyncio.run(RequestBodyLimit(downstream, max_bytes=max_bytes)(scope, receive, send))
        return status, called


if __name__ == "__main__":
    unittest.main()
