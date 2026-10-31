"""Regression tests for HTTP token parsing and bounded terminal ASGI replay."""
import unittest

from http_body_limit import MAX_EMPTY_REQUEST_EVENTS, RequestBodyLimit


class RequestBodyTerminalSafetyTests(unittest.IsolatedAsyncioTestCase):
    async def invoke(self, frames, headers=(), method="POST", http_version="1.1",
                     inner=None):
        dispatched = False

        async def default_inner(scope, receive, send):
            return await receive()

        async def wrapped_inner(scope, receive, send):
            nonlocal dispatched
            dispatched = True
            return await (inner or default_inner)(scope, receive, send)

        iterator = iter(frames)

        async def receive():
            return next(iterator)

        sent = []

        async def send(message):
            sent.append(message)

        scope = {"type": "http", "method": method, "headers": list(headers),
                 "http_version": http_version}
        result = await RequestBodyLimit(wrapped_inner, max_bytes=12)(
            scope, receive, send)
        return (sent[0]["status"] if sent else 200, dispatched, result)

    async def test_invalid_method_tokens_do_not_bypass_body_validation(self):
        for method in (" POST", "POST ", "PO/ST", "POST\n", "PO\u017fT",
                       "", "POST\t", "\u212aET", b"POST"):
            with self.subTest(method=method):
                status, dispatched, _ = await self.invoke([], method=method)
                self.assertEqual(status, 400)
                self.assertFalse(dispatched)

    async def test_mixed_case_post_is_rejected_as_distinct_method(self):
        # RFC 9110 methods are case-sensitive; mixed-case variants must not
        # bypass the POST body guard by falling through to the inner app.
        for method in ("PoSt", "post", "POST".lower()):
            with self.subTest(method=method):
                status, dispatched, _ = await self.invoke(
                    [{"type": "http.request", "body": b"ok"}], method=method)
                self.assertEqual(status, 400)
                self.assertFalse(dispatched)

    async def test_exact_post_remains_bounded(self):
        status, dispatched, result = await self.invoke(
            [{"type": "http.request", "body": b"ok"}], method="POST")
        self.assertEqual(status, 200)
        self.assertTrue(dispatched)
        self.assertEqual(result["body"], b"ok")
        status, dispatched, _ = await self.invoke(
            [{"type": "http.request", "body": b"x" * 13}], method="POST")
        self.assertEqual(status, 413)
        self.assertFalse(dispatched)

    async def test_rejects_control_bytes_in_header_values(self):
        for value in (1, 2, 8, 11, 12, 31, 127):
            with self.subTest(value=value):
                status, dispatched, _ = await self.invoke(
                    [], headers=[(b"x-test", b"safe" + bytes([value]))])
                self.assertEqual(status, 400)
                self.assertFalse(dispatched)

    async def test_preserves_legal_horizontal_tab_and_obs_text(self):
        status, dispatched, result = await self.invoke(
            [{"type": "http.request", "body": b"ok"}],
            headers=[(b"x-test", b"one\ttwo\xff")])
        self.assertEqual(status, 200)
        self.assertTrue(dispatched)
        self.assertEqual(result["body"], b"ok")

    async def test_chunked_transfer_encoding_requires_http_1_1(self):
        for version in ("1.0", "2", "3"):
            with self.subTest(version=version):
                status, dispatched, _ = await self.invoke(
                    [], headers=[(b"transfer-encoding", b"chunked")],
                    http_version=version)
                self.assertEqual(status, 400)
                self.assertFalse(dispatched)

    async def test_chunked_transfer_encoding_remains_valid_in_http_1_1(self):
        status, dispatched, result = await self.invoke(
            [{"type": "http.request", "body": b"ok"}],
            headers=[(b"transfer-encoding", b"chunked")])
        self.assertEqual(status, 200)
        self.assertTrue(dispatched)
        self.assertEqual(result["body"], b"ok")

    async def test_rejects_excessive_zero_progress_chunks(self):
        frames = ({"type": "http.request", "body": b"", "more_body": True}
                  for _ in range(MAX_EMPTY_REQUEST_EVENTS + 1))
        status, dispatched, _ = await self.invoke(frames)
        self.assertEqual(status, 413)
        self.assertFalse(dispatched)

    async def test_allows_empty_chunks_at_boundary_then_final_body(self):
        frames = ({"type": "http.request", "body": b"", "more_body": True}
                  for _ in range(MAX_EMPTY_REQUEST_EVENTS))
        def combined():
            yield from frames
            yield {"type": "http.request", "body": b"ok", "more_body": False}
        status, dispatched, result = await self.invoke(combined())
        self.assertEqual(status, 200)
        self.assertTrue(dispatched)
        self.assertEqual(result["body"], b"ok")

    async def test_post_terminal_request_frames_are_not_forwarded(self):
        async def read_twice(scope, receive, send):
            return await receive(), await receive()
        frames = [
            {"type": "http.request", "body": b"ok", "more_body": False},
            {"type": "http.request", "body": b"unbounded", "more_body": False},
        ]
        status, dispatched, result = await self.invoke(frames, inner=read_twice)
        self.assertEqual(status, 200)
        self.assertTrue(dispatched)
        self.assertEqual(result[0]["body"], b"ok")
        self.assertEqual(result[1], {"type": "http.disconnect"})

    async def test_post_terminal_disconnect_is_forwarded(self):
        async def read_twice(scope, receive, send):
            return await receive(), await receive()
        frames = [
            {"type": "http.request", "body": b"ok", "more_body": False},
            {"type": "http.disconnect"},
        ]
        status, dispatched, result = await self.invoke(frames, inner=read_twice)
        self.assertEqual(status, 200)
        self.assertTrue(dispatched)
        self.assertEqual(result[1], {"type": "http.disconnect"})


if __name__ == "__main__":
    unittest.main()
