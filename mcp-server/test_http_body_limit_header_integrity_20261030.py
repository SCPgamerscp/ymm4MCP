"""Regression tests for malformed ASGI HTTP header envelopes and framing."""
import unittest

from http_body_limit import MAX_HEADER_BYTES, MAX_HEADER_FIELDS, RequestBodyLimit


class HeaderIntegrityTests(unittest.IsolatedAsyncioTestCase):
    async def invoke(self, headers, frames=None, method="POST"):
        dispatched = False

        async def inner(scope, receive, send):
            nonlocal dispatched
            dispatched = True
            return await receive()

        iterator = iter(frames if frames is not None else [])
        async def receive():
            return next(iterator)

        messages = []
        async def send(message):
            messages.append(message)

        app = RequestBodyLimit(inner, max_bytes=12)
        result = await app({"type": "http", "method": method, "headers": headers},
                           receive, send)
        return (messages[0]["status"] if messages else 200, dispatched, result)

    async def test_rejects_invalid_header_containers(self):
        for headers in (None, {}, b"content-length", "bad"):
            with self.subTest(headers=headers):
                status, dispatched, _ = await self.invoke(headers)
                self.assertEqual(status, 400)
                self.assertFalse(dispatched)

    async def test_rejects_malformed_header_entries(self):
        for field in (b"bad", (), (b"host",), (b"host", b"x", b"y"), None):
            with self.subTest(field=field):
                status, dispatched, _ = await self.invoke([field])
                self.assertEqual(status, 400)
                self.assertFalse(dispatched)

    async def test_rejects_nonbyte_header_name_and_value(self):
        for field in (("content-length", b"3"), (b"content-length", "3"),
                      (None, b"3"), (b"host", None)):
            with self.subTest(field=field):
                status, dispatched, _ = await self.invoke([field])
                self.assertEqual(status, 400)
                self.assertFalse(dispatched)

    async def test_rejects_invalid_header_field_names(self):
        for name in (b"", b"bad name", b"bad:name", b"bad\rname",
                     b"bad\nname", b"bad\x00name", b"bad\xffname"):
            with self.subTest(name=name):
                status, dispatched, _ = await self.invoke([(name, b"ok")])
                self.assertEqual(status, 400)
                self.assertFalse(dispatched)

    async def test_rejects_header_value_injection(self):
        for value in (b"a\rInjected: yes", b"a\nInjected: yes", b"a\x00b"):
            with self.subTest(value=value):
                status, dispatched, _ = await self.invoke([(b"x-test", value)])
                self.assertEqual(status, 400)
                self.assertFalse(dispatched)

    async def test_rejects_too_many_header_fields(self):
        status, dispatched, _ = await self.invoke([(b"x", b"") for _ in range(MAX_HEADER_FIELDS + 1)])
        self.assertEqual(status, 400)
        self.assertFalse(dispatched)

    async def test_rejects_excessive_total_header_bytes(self):
        status, dispatched, _ = await self.invoke([(b"x", b"a" * MAX_HEADER_BYTES)])
        self.assertEqual(status, 431)
        self.assertFalse(dispatched)

    async def test_rejects_ambiguous_transfer_encodings(self):
        for headers in (
            [(b"transfer-encoding", b"gzip")],
            [(b"transfer-encoding", b"gzip, chunked")],
            [(b"transfer-encoding", b"chunked, chunked")],
            [(b"transfer-encoding", b"chunked"), (b"Transfer-Encoding", b"chunked")],
            [(b"transfer-encoding", b"")],
        ):
            with self.subTest(headers=headers):
                status, dispatched, _ = await self.invoke(headers)
                self.assertEqual(status, 400)
                self.assertFalse(dispatched)

    async def test_accepts_valid_chunked_encoding_with_ows(self):
        status, dispatched, result = await self.invoke(
            [(b"TrAnSfEr-EnCoDiNg", b"  ChUnKeD\t")],
            [{"type": "http.request", "body": b"ab", "more_body": True},
             {"type": "http.request", "body": b"cd", "more_body": False}])
        self.assertEqual(status, 200)
        self.assertTrue(dispatched)
        self.assertEqual(result["body"], b"abcd")

    async def test_accepts_nonascii_header_values_and_mixed_case_names(self):
        status, dispatched, result = await self.invoke(
            [(b"X-Label", b"\xff\xfe"), (b"CoNtEnT-LeNgTh", b"2")],
            [{"type": "http.request", "body": b"ok", "more_body": False}])
        self.assertEqual(status, 200)
        self.assertTrue(dispatched)
        self.assertEqual(result["body"], b"ok")

    async def test_rejects_nonstring_method_without_dispatch(self):
        for method in (None, b"POST", 123):
            with self.subTest(method=method):
                status, dispatched, _ = await self.invoke([], method=method)
                self.assertEqual(status, 400)
                self.assertFalse(dispatched)


if __name__ == "__main__":
    unittest.main()
