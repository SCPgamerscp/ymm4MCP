"""IPv6 spelling aliases must not duplicate HTTP Host allowlist rules (Issue #22)."""
import unittest

from http_host_validation import parse_allowed_hosts


class IPv6AliasTests(unittest.TestCase):
    def test_rejects_equivalent_ipv6_hosts(self):
        aliases = (
            ("[::1]", "[0:0:0:0:0:0:0:1]"),
            ("[2001:db8::1]", "[2001:0DB8:0:0:0:0:0:1]"),
            ("[::ffff:192.0.2.1]", "[0:0:0:0:0:ffff:c000:201]"),
        )
        for first, second in aliases:
            with self.subTest(first=first, second=second):
                with self.assertRaisesRegex(ValueError, "duplicate"):
                    parse_allowed_hosts(f"{first},{second}")

    def test_rejects_equivalent_ipv6_hosts_with_same_port(self):
        with self.assertRaisesRegex(ValueError, "duplicate"):
            parse_allowed_hosts("[2001:db8::1]:443,[2001:0DB8:0:0:0:0:0:1]:443")

    def test_distinct_ports_remain_valid(self):
        hosts = "[::1],[0:0:0:0:0:0:0:1]:443,[::1]:*"
        self.assertEqual(parse_allowed_hosts(hosts), hosts.split(","))

    def test_existing_ipv4_dns_and_ipv6_host_rules_still_work(self):
        hosts = "localhost,LOCALHOST:*,127.0.0.1,[::1]:65535"
        self.assertEqual(parse_allowed_hosts(hosts), hosts.split(","))


if __name__ == "__main__":
    unittest.main()
