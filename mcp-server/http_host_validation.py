"""Validate explicit Host-header allowlist entries for the remote MCP transport."""

import ipaddress
import re

_DNS_LABEL = re.compile(r"[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?\Z")


def _valid_host(host: str) -> bool:
    if not host or len(host) > 253 or any(ord(c) > 127 for c in host):
        return False
    if host.startswith("[") and host.endswith("]"):
        try:
            ipaddress.IPv6Address(host[1:-1])
        except ValueError:
            return False
        return True
    if ":" in host or "[" in host or "]" in host:
        return False
    if all(part.isascii() and part.isdigit() for part in host.split(".")) and "." in host:
        try:
            ipaddress.IPv4Address(host)
        except ValueError:
            return False
        return True
    return all(_DNS_LABEL.fullmatch(label) for label in host.split("."))


def parse_allowed_hosts(raw: str) -> list[str]:
    """Accept explicit DNS/IP hosts, optionally followed by a port or :*.

    Do not accept wildcard hostnames, invalid ports, URL syntax, or ambiguous
    comma-separated entries. Bracketed IPv6 literals are supported.
    """
    if not isinstance(raw, str) or any(ord(ch) < 32 or ord(ch) == 127 for ch in raw):
        raise ValueError("YMM4_MCP_ALLOWED_HOSTS must list explicit hosts")
    hosts = [entry.strip(" ") for entry in raw.split(",")]
    if not hosts or any(not entry for entry in hosts):
        raise ValueError("YMM4_MCP_ALLOWED_HOSTS must list explicit hosts")
    seen = set()
    for entry in hosts:
        if any(ch.isspace() or ord(ch) > 126 for ch in entry):
            raise ValueError("YMM4_MCP_ALLOWED_HOSTS must list explicit hosts")
        if entry.startswith("["):
            close = entry.find("]")
            if close == -1:
                raise ValueError("YMM4_MCP_ALLOWED_HOSTS has an invalid IPv6 host")
            host, suffix = entry[:close + 1], entry[close + 1:]
            if suffix and not suffix.startswith(":"):
                raise ValueError("YMM4_MCP_ALLOWED_HOSTS has an invalid port")
            port = suffix[1:] if suffix else None
        else:
            parts = entry.split(":")
            if len(parts) > 2:
                raise ValueError("YMM4_MCP_ALLOWED_HOSTS requires bracketed IPv6 hosts")
            host, port = parts[0], parts[1] if len(parts) == 2 else None
        if not _valid_host(host):
            raise ValueError("YMM4_MCP_ALLOWED_HOSTS must list explicit hosts")
        if port is not None and port != "*":
            if (not port.isascii() or not port.isdecimal() or
                    (len(port) > 1 and port[0] == "0") or
                    not 1 <= int(port) <= 65535):
                raise ValueError("YMM4_MCP_ALLOWED_HOSTS has an invalid port")
        canonical = entry.lower()
        if canonical in seen:
            raise ValueError("YMM4_MCP_ALLOWED_HOSTS has duplicate entries")
        seen.add(canonical)
    return hosts
