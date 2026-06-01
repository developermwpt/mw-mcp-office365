"""Fixtures partilhadas de teste (QA).

Tudo com Graph/Entra **mockados** — nenhum teste toca no tenant real, na rede ou dorme de
verdade. O relógio é fixo e injetável; o `sleeper` do GraphClient é um no-op assíncrono.
"""

from __future__ import annotations

import base64
import os
from datetime import datetime, timedelta, timezone

import pytest

from mcp_o365.config import Settings
from mcp_o365.identity.mapping import IdentityMapping
from mcp_o365.storage.crypto import LocalAesGcmCipher
from mcp_o365.storage.token_store import TokenStore

FIXED_NOW = datetime(2026, 6, 1, 12, 0, 0, tzinfo=timezone.utc)


class Clock:
    """Relógio controlável para testes."""

    def __init__(self, now: datetime = FIXED_NOW) -> None:
        self._now = now

    def __call__(self) -> datetime:
        return self._now

    def advance(self, seconds: float) -> None:
        self._now = self._now + timedelta(seconds=seconds)


@pytest.fixture
def clock() -> Clock:
    return Clock()


@pytest.fixture
def cipher() -> LocalAesGcmCipher:
    return LocalAesGcmCipher(b"\x00" * 32)


@pytest.fixture
def store(cipher: LocalAesGcmCipher, clock: Clock) -> TokenStore:
    s = TokenStore(":memory:", cipher, clock=clock)
    yield s
    s.close()


@pytest.fixture
def mapping(store: TokenStore, clock: Clock) -> IdentityMapping:
    return IdentityMapping(store, clock=clock)


@pytest.fixture
def config() -> Settings:
    """Configuração de teste (valores fake, sem segredos reais)."""
    env = {
        "ENTRA_TENANT_ID": "tenant-test",
        "ENTRA_CLIENT_ID": "client-test",
        "ENTRA_CLIENT_SECRET": "fake-secret",
        "ENTRA_AUTHORITY": "https://login.microsoftonline.com/tenant-test",
        "OAUTH_REDIRECT_URI": "https://mcp.example.com/callback",
        "GRAPH_SCOPES": "User.Read offline_access openid profile",
        "MCP_ISSUER_URL": "https://mcp.example.com",
        "MCP_PUBLIC_BASE_URL": "https://mcp.example.com",
        "TOKEN_STORE_PATH": ":memory:",
        "TOKEN_ENCRYPTION_KEY": base64.b64encode(b"\x00" * 32).decode(),
        "LOG_LEVEL": "INFO",
        "BIND_HOST": "127.0.0.1",
        "BIND_PORT": "8000",
    }
    old = dict(os.environ)
    os.environ.update(env)
    try:
        yield Settings(_env_file=None)  # type: ignore[call-arg]
    finally:
        os.environ.clear()
        os.environ.update(old)


class FakeMsalApp:
    """Fake da ConfidentialClientApplication do msal, com respostas programáveis."""

    def __init__(
        self,
        *,
        code_result: dict | None = None,
        refresh_result: dict | None = None,
        auth_url: str = "https://login.microsoftonline.com/authorize?fake=1",
    ) -> None:
        self._code_result = code_result or {}
        self._refresh_result = refresh_result or {}
        self._auth_url = auth_url
        self.calls: list[str] = []

    def get_authorization_request_url(self, scopes, state=None, redirect_uri=None):
        self.calls.append("authorize")
        return f"{self._auth_url}&state={state}"

    def acquire_token_by_authorization_code(self, code, scopes=None, redirect_uri=None):
        self.calls.append("exchange_code")
        return self._code_result

    def acquire_token_by_refresh_token(self, refresh_token, scopes=None):
        self.calls.append("refresh")
        return self._refresh_result


def graph_token_response(
    *, refresh_token: str = "rt-1", expires_in: int = 3600, oid: str = "user-oid-1"
) -> dict:
    """Resposta de sucesso canónica do Entra (token Graph)."""
    return {
        "access_token": "graph-access-1",
        "refresh_token": refresh_token,
        "expires_in": expires_in,
        "scope": "User.Read offline_access openid profile",
        "id_token_claims": {
            "oid": oid,
            "tid": "tenant-test",
            "preferred_username": "user@example.com",
        },
    }


INVALID_GRANT_RESPONSE = {
    "error": "invalid_grant",
    "error_description": "AADSTS50173: refresh token expired/revoked or blocked by policy.",
}

CONSENT_REQUIRED_RESPONSE = {
    "error": "consent_required",
    "error_description": "AADSTS65001: consent required.",
}
