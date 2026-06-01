"""Integração — health e metadata OAuth (T7/T12) servidos pelo SDK.

Usa o TestClient da Starlette sobre a app ASGI do FastMCP. Não toca em Entra/Graph.
"""

from __future__ import annotations

from starlette.testclient import TestClient

from mcp_o365.app import build_components


def _client(config) -> TestClient:
    comp = build_components(config)
    return TestClient(comp["server"].streamable_http_app())


def test_healthz(config):
    with _client(config) as c:
        r = c.get("/healthz")
        assert r.status_code == 200
        assert r.json()["status"] == "ok"


def test_readyz(config):
    with _client(config) as c:
        r = c.get("/readyz")
        assert r.status_code == 200
        body = r.json()
        assert body["status"] == "ready"
        assert body["db"] == "ok"


def test_well_known_protected_resource(config):
    # RFC 9728
    with _client(config) as c:
        r = c.get("/.well-known/oauth-protected-resource")
        assert r.status_code == 200
        data = r.json()
        assert "resource" in data
        assert "authorization_servers" in data


def test_well_known_authorization_server(config):
    # RFC 8414 — inclui registration_endpoint (DCR, RFC 7591).
    with _client(config) as c:
        r = c.get("/.well-known/oauth-authorization-server")
        assert r.status_code == 200
        data = r.json()
        assert data["issuer"].rstrip("/") == "https://mcp.example.com"
        assert "authorization_endpoint" in data
        assert "token_endpoint" in data
        assert "registration_endpoint" in data


def test_callback_sem_parametros_da_400(config):
    with _client(config) as c:
        r = c.get("/callback")
        assert r.status_code == 400


def test_tool_requer_autenticacao(config):
    # Sem token do Plano A, o transporte MCP recusa (401).
    with _client(config) as c:
        r = c.post("/mcp", json={"jsonrpc": "2.0", "id": 1, "method": "tools/list"})
        assert r.status_code == 401
