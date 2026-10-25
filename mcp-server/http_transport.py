"""Optional TLS-only Streamable HTTP transport for the existing MCP server.

The bearer credential here is intentionally separate from the local plugin token.
No unauthenticated fallback, wildcard CORS, proxy trust, or automatic public hosting.
"""
from contextlib import asynccontextmanager
import hmac
import os

from mcp.server.streamable_http_manager import StreamableHTTPSessionManager
from mcp.server.transport_security import TransportSecuritySettings
from starlette.applications import Starlette
from starlette.responses import JSONResponse
from starlette.routing import Route

from http_host_validation import parse_allowed_hosts


def create_http_app(mcp_server, close_client):
    token = os.environ.get("YMM4_MCP_BEARER_TOKEN", "")
    if len(token) < 32 or any(not 33 <= ord(c) <= 126 for c in token):
        raise ValueError("YMM4_MCP_BEARER_TOKEN must be a distinct, random ASCII token of at least 32 characters")
    hosts = parse_allowed_hosts(os.environ.get(
        "YMM4_MCP_ALLOWED_HOSTS", "127.0.0.1,127.0.0.1:*,localhost,localhost:*"
    ))
    manager = StreamableHTTPSessionManager(
        app=mcp_server, json_response=True, stateless=True,
        security_settings=TransportSecuritySettings(
            enable_dns_rebinding_protection=True, allowed_hosts=hosts, allowed_origins=[]
        ),
    )

    @asynccontextmanager
    async def lifespan(app):
        try:
            async with manager.run():
                yield
        finally:
            await close_client()

    class Endpoint:
        async def __call__(self, scope, receive, send):
            await manager.handle_request(scope, receive, send)

    application = Starlette(routes=[Route("/mcp", Endpoint(), methods=["GET", "POST", "DELETE"])], lifespan=lifespan)

    class AuthenticatedApplication:
        async def __call__(self, scope, receive, send):
            if scope["type"] == "http":
                # Never collapse repeated Authorization headers: different proxies
                # may select different values, making the authentication ambiguous.
                auth_values = [value for name, value in scope["headers"]
                               if name.lower() == b"authorization"]
                expected = ("Bearer " + token).encode("ascii")
                if len(auth_values) != 1 or not hmac.compare_digest(auth_values[0], expected):
                    await JSONResponse({"error": "Unauthorized"}, status_code=401,
                                       headers={"WWW-Authenticate": "Bearer"})(scope, receive, send)
                    return
                if (scope.get("scheme") != "https" or
                        any(name.lower() == b"origin" for name, _ in scope["headers"])):
                    await JSONResponse({"error": "TLS required; browser origins are not supported"}, status_code=403)(scope, receive, send)
                    return
            await application(scope, receive, send)

    return AuthenticatedApplication()


async def run_http(mcp_server, close_client, args):
    import uvicorn
    if not args.tls_cert or not args.tls_key:
        raise ValueError("Streamable HTTP requires --tls-cert and --tls-key; plain HTTP is not supported")
    application = create_http_app(mcp_server, close_client)
    config = uvicorn.Config(application, host=args.host, port=args.port,
                            ssl_certfile=args.tls_cert, ssl_keyfile=args.tls_key,
                            proxy_headers=False, access_log=False)
    await uvicorn.Server(config).serve()
