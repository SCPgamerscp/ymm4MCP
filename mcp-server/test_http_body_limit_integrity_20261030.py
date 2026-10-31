"""Regression coverage for malformed or abusive remote MCP body streams."""
import unittest

from http_body_limit import MAX_REQUEST_EVENTS, RequestBodyLimit


class RequestBodyIntegrityTests(unittest.IsolatedAsyncioTestCase):
    async def invoke(self, frames, headers=(), limit=12):
        dispatched = False
        async def inner(scope, receive, send):
            nonlocal dispatched
            dispatched = True
            return await receive()
        app = RequestBodyLimit(inner, max_bytes=limit)
        iterator = iter(frames)
        async def receive():
            return next(iterator)
        messages = []
        async def send(message):
            messages.append(message)
        result = await app({"type": "http", "method": "POST", "headers": list(headers)}, receive, send)
        return (messages[0]["status"] if messages else 200, dispatched, result)

    async def test_rejects_noncanonical_content_length_without_reading(self):
        for value in (b"00", b"012", b"00000000000000000000"):
            with self.subTest(value=value):
                status, dispatched, _ = await self.invoke([], [(b"content-length", value)])
                self.assertEqual(status, 400)
                self.assertFalse(dispatched)

    async def test_rejects_transfer_encoding_and_content_length_conflict(self):
        status, dispatched, _ = await self.invoke([], [
            (b"content-length", b"4"), (b"transfer-encoding", b"chunked")])
        self.assertEqual(status, 400)
        self.assertFalse(dispatched)

    async def test_rejects_nonbyte_content_length(self):
        for value in ("2", 2, None):
            with self.subTest(value=value):
                status, dispatched, _ = await self.invoke([], [(b"content-length", value)])
                self.assertEqual(status, 400)
                self.assertFalse(dispatched)

    async def test_rejects_nonbyte_request_bodies(self):
        for value in ("abc", bytearray(b"abc"), None, 12):
            with self.subTest(value=value):
                status, dispatched, _ = await self.invoke([{"type": "http.request", "body": value}])
                self.assertEqual(status, 400)
                self.assertFalse(dispatched)

    async def test_rejects_nonboolean_more_body(self):
        for value in (1, "false", None):
            with self.subTest(value=value):
                status, dispatched, _ = await self.invoke([{"type": "http.request", "body": b"a", "more_body": value}])
                self.assertEqual(status, 400)
                self.assertFalse(dispatched)

    async def test_rejects_malformed_receive_envelopes(self):
        for value in (None, b"abc", [], {"body": b"x"}):
            with self.subTest(value=value):
                status, dispatched, _ = await self.invoke([value])
                self.assertEqual(status, 400)
                self.assertFalse(dispatched)

    async def test_declared_length_overflow_stops_at_first_offending_chunk(self):
        status, dispatched, _ = await self.invoke([
            {"type": "http.request", "body": b"abc", "more_body": True},
            {"type": "http.request", "body": b"d", "more_body": True},
        ], [(b"content-length", b"3")])
        self.assertEqual(status, 400)
        self.assertFalse(dispatched)

    async def test_bounded_event_count_for_zero_length_chunk_flood(self):
        frames = ({"type": "http.request", "body": b"", "more_body": True}
                  for _ in range(MAX_REQUEST_EVENTS + 1))
        status, dispatched, _ = await self.invoke(frames)
        self.assertEqual(status, 413)
        self.assertFalse(dispatched)

    async def test_accepts_valid_multichunk_body_with_exact_length(self):
        status, dispatched, result = await self.invoke([
            {"type": "http.request", "body": b"ab", "more_body": True},
            {"type": "http.request", "body": b"cd", "more_body": False},
        ], [(b"content-length", b"4")])
        self.assertEqual(status, 200)
        self.assertTrue(dispatched)
        self.assertEqual(result["body"], b"abcd")

    async def test_accepts_valid_empty_body_with_zero_content_length(self):
        status, dispatched, result = await self.invoke([
            {"type": "http.request", "body": b"", "more_body": False},
        ], [(b"content-length", b"0")])
        self.assertEqual(status, 200)
        self.assertTrue(dispatched)
        self.assertEqual(result["body"], b"")


if __name__ == "__main__":
    unittest.main()
