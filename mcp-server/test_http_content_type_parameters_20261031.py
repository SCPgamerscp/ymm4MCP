"""Regression cases for Content-Type parameter parsing."""
import importlib.util
from pathlib import Path
import unittest

spec = importlib.util.spec_from_file_location("http_body_limit", Path(__file__).with_name("http_body_limit.py"))
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)


class ContentTypeParameterTests(unittest.TestCase):
    def test_valid_parameters(self):
        for value in (
            b"application/json",
            b"application/json; charset=utf-8",
            b'application/json; note="hello; world"',
            b'application/json; note="a\\\"b"; charset=utf-8',
            b'application/json; empty=""',
        ):
            with self.subTest(value=value):
                self.assertFalse(module._invalid_content_type_parameters(value))

    def test_invalid_parameters(self):
        for value in (
            b"application/json;",
            b"application/json; charset",
            b"application/json; =utf-8",
            b"application/json; charset=",
            b"application/json; charset=utf-8;",
            b"application/json; charset=utf-8;;foo=bar",
            b"application/json; char set=utf-8",
            b"application/json; charset=hello world",
            b'application/json; note="unclosed',
            b'application/json; note="a"junk',
        ):
            with self.subTest(value=value):
                self.assertTrue(module._invalid_content_type_parameters(value))


if __name__ == "__main__":
    unittest.main()
