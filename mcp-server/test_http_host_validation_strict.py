"""Regression tests for strict MCP Host-header allowlist validation (Issue #22)."""
import unittest

from http_host_validation import parse_allowed_hosts


class StrictHostValidationTests(unittest.TestCase):
    def test_valid_dns_ipv4_ipv6_and_port_boundaries(self):
        value = "api.example.com:1,api.example.com:65535,127.0.0.1:*,[::1],[2001:db8::1]:443"
        self.assertEqual(parse_allowed_hosts(value), value.split(","))

    def test_rejects_malformed_or_out_of_range_ports(self):
        for value in ("localhost:", "localhost:0", "localhost:65536", "localhost:-1",
                      "localhost:+1", "localhost:abc", "localhost:1:2",
                      "localhost:01", "localhost:1.0", "localhost: *x"):
            with self.subTest(value=value), self.assertRaises(ValueError):
                parse_allowed_hosts(value)

    def test_rejects_invalid_dns_and_wildcards(self):
        for value in ("-host.example", "host-.example", "a..b", ".example",
                      "example.", "bad_host", "a" * 64 + ".example",
                      "*.example.com", "api*.example.com", "http://example.com",
                      "user@example.com", "foo/bar", "foo\\bar"):
            with self.subTest(value=value), self.assertRaises(ValueError):
                parse_allowed_hosts(value)

    def test_rejects_invalid_numeric_ipv4_aliases(self):
        for value in ("256.1.1.1", "127.0.0.999", "01.2.3.4", "1.2.3"):
            with self.subTest(value=value), self.assertRaises(ValueError):
                parse_allowed_hosts(value)

    def test_ipv6_must_be_bracketed_and_valid(self):
        for value in ("::1", "[::1", "::1]", "[not-ipv6]", "[::1]bad",
                      "[::1]:0", "[::1]:65536", "[::1]:-1", "[fe80::1%eth0]"):
            with self.subTest(value=value), self.assertRaises(ValueError):
                parse_allowed_hosts(value)

    def test_empty_entries_are_not_silently_ignored(self):
        for value in ("", ",localhost", "localhost,", "localhost,,127.0.0.1", " , "):
            with self.subTest(value=value), self.assertRaises(ValueError):
                parse_allowed_hosts(value)

    def test_duplicate_hosts_case_insensitive(self):
        for value in ("localhost,LOCALHOST", "api.example.com:*,API.EXAMPLE.COM:*"):
            with self.subTest(value=value), self.assertRaises(ValueError):
                parse_allowed_hosts(value)


if __name__ == "__main__":
    unittest.main()
