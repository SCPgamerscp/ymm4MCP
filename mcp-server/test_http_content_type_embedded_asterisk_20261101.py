"""Regression coverage for literal asterisks in Content-Type tokens."""
import unittest
from http_body_limit import _invalid_media_type


class ContentTypeAsteriskTests(unittest.TestCase):
    def test_embedded_asterisk_is_valid_token_character(self):
        for value in (b"application/vnd.foo*bar", b"x*y/z", b"x/y*z"):
            with self.subTest(value=value):
                self.assertFalse(_invalid_media_type(value))

    def test_wildcard_media_ranges_still_rejected(self):
        for value in (b"*/*", b"application/*", b"*/json"):
            with self.subTest(value=value):
                self.assertTrue(_invalid_media_type(value))

    def test_other_invalid_tokens_still_rejected(self):
        for value in (b"application/", b"/json", b"application/j son"):
            with self.subTest(value=value):
                self.assertTrue(_invalid_media_type(value))


if __name__ == "__main__":
    unittest.main()
