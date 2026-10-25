"""Regression tests for explicit HTTP Host allowlisting (Issue #22)."""
import unittest

from http_host_validation import parse_allowed_hosts


class HostValidationTests(unittest.TestCase):
    def test_default_and_explicit_hosts(self):
        self.assertEqual(parse_allowed_hosts("127.0.0.1,127.0.0.1:*,localhost,localhost:*"),
                         ["127.0.0.1", "127.0.0.1:*", "localhost", "localhost:*"])
        self.assertEqual(parse_allowed_hosts(" api.example.com:443 , api.example.com:* "),
                         ["api.example.com:443", "api.example.com:*"])

    def test_rejects_wildcard_hosts_and_invalid_patterns(self):
        for raw in ("*", "*:*", "*.example.com", "*.example.com:*", "api*.example.com",
                    "api.example.com*", ":*", "api.example.com/**", "api.example.com@evil.com",
                    "api.example.com\\evil.com", "api.example.com?x", "api.example.com#x"):
            with self.subTest(raw=raw), self.assertRaises(ValueError):
                parse_allowed_hosts(raw)

    def test_rejects_header_controls_and_whitespace(self):
        for raw in ("", " , ", "localhost\n", "localhost\r\nexample.com",
                    "localhost\t", "local host", "localhost\x7f", "localhost\x00"):
            with self.subTest(raw=repr(raw)), self.assertRaises(ValueError):
                parse_allowed_hosts(raw)


if __name__ == "__main__":
    unittest.main()
