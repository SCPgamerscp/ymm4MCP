"""Regression tests for rejecting Accept-style wildcards in Content-Type."""
import importlib.util
from pathlib import Path
import unittest

MODULE_PATH = Path(__file__).with_name("http_body_limit.py")
spec = importlib.util.spec_from_file_location("http_body_limit", MODULE_PATH)
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)


class ContentTypeWildcardTests(unittest.TestCase):
    def test_rejects_wildcard_media_ranges(self):
        for value in (b"*/*", b"application/*", b"*/json",
                      b"application/j*son", b"app*lication/json",
                      b"application/*; charset=utf-8"):
            with self.subTest(value=value):
                self.assertTrue(module._invalid_media_type(value))

    def test_allows_concrete_media_types(self):
        for value in (b"application/json", b"text/plain; charset=utf-8",
                      b"application/vnd.example+json", b"application/x.custom"):
            with self.subTest(value=value):
                self.assertFalse(module._invalid_media_type(value))


if __name__ == "__main__":
    unittest.main()
