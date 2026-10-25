"""Validate explicit Host-header allowlist entries for the remote MCP transport."""


def parse_allowed_hosts(raw: str) -> list[str]:
    """Allow named hosts with an optional wildcard *port*, never wildcard hosts.

    The MCP transport's DNS-rebinding guard accepts wildcard host patterns.
    Here we deliberately support only `hostname` and `hostname:*` entries.
    """
    if not isinstance(raw, str) or any(ord(ch) < 32 or ord(ch) == 127 for ch in raw):
        raise ValueError("YMM4_MCP_ALLOWED_HOSTS must list explicit hosts")
    hosts = [entry.strip(" ") for entry in raw.split(",") if entry.strip(" ")]
    if not hosts:
        raise ValueError("YMM4_MCP_ALLOWED_HOSTS must list explicit hosts")
    for host in hosts:
        if (any(ch.isspace() or ord(ch) > 126 for ch in host)
                or any(ch in host for ch in "/\\@?#")
                or ("*" in host and (not host.endswith(":*") or host.count("*") != 1))
                or host in {"*", "*:*", ":*"}):
            raise ValueError("YMM4_MCP_ALLOWED_HOSTS must list explicit hosts")
    return hosts
