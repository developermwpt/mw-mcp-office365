"""Integração — fluxo OAuth do Plano A ao nível do provider (T9).

Exercita a lógica própria (DCR, authorize, callback do Entra, authorization code e access
token) sem depender do transporte HTTP do SDK. O Entra é substituído por um FakeMsalApp.
"""

from __future__ import annotations

import pytest
from mcp.server.auth.provider import AuthorizationParams
from mcp.shared.auth import OAuthClientInformationFull

from mcp_o365.auth.plane_a import MwOAuthProvider
from mcp_o365.auth.plane_b import PlaneB
from tests.conftest import FakeMsalApp, graph_token_response


@pytest.fixture
def provider(store, mapping, config, clock):
    fake = FakeMsalApp(code_result=graph_token_response(oid="subject-xyz"))
    plane_b = PlaneB(config, msal_app_factory=lambda: fake, clock=clock)
    return MwOAuthProvider(
        store=store, plane_b=plane_b, mapping=mapping, config=config, clock=clock
    )


def _client() -> OAuthClientInformationFull:
    return OAuthClientInformationFull(
        client_id="claude-client-1",
        redirect_uris=["https://claude.ai/callback"],
        token_endpoint_auth_method="none",
    )


async def test_dcr_persiste_cliente(provider):
    client = _client()
    await provider.register_client(client)
    got = await provider.get_client("claude-client-1")
    assert got is not None
    assert str(got.redirect_uris[0]) == "https://claude.ai/callback"


async def test_fluxo_completo_authorize_callback_token(provider):
    client = _client()
    await provider.register_client(client)

    # 1) authorize -> redireciona para o Entra e guarda a transação (state).
    params = AuthorizationParams(
        state="claude-state",
        scopes=["User.Read"],
        code_challenge="challenge-123",
        redirect_uri="https://claude.ai/callback",
        redirect_uri_provided_explicitly=True,
    )
    entra_url = await provider.authorize(client, params)
    assert entra_url.startswith("https://login.microsoftonline.com/")
    state = entra_url.split("state=")[1]

    # 2) callback do Entra -> liga a conta e emite o authorization code do Plano A.
    redirect_back = provider.complete_entra_callback(code="entra-code", state=state)
    assert redirect_back.startswith("https://claude.ai/callback")
    assert "state=claude-state" in redirect_back
    plana_code = redirect_back.split("code=")[1].split("&")[0]

    # 3) load + exchange do authorization code -> access token MCP.
    loaded = await provider.load_authorization_code(client, plana_code)
    assert loaded is not None
    assert loaded.subject == "subject-xyz"
    token = await provider.exchange_authorization_code(client, loaded)
    assert token.access_token

    # 4) o access token resolve para o subject correto.
    resolved = await provider.load_access_token(token.access_token)
    assert resolved is not None
    assert resolved.subject == "subject-xyz"

    # 5) a conta Graph ficou ligada ao subject (mapping).
    session = provider._mapping.get_session("subject-xyz")
    assert session is not None
    assert session.default_account.access_token == "graph-access-1"


async def test_authorization_code_so_usado_uma_vez(provider):
    client = _client()
    await provider.register_client(client)
    params = AuthorizationParams(
        state="s", scopes=["User.Read"], code_challenge="c",
        redirect_uri="https://claude.ai/callback", redirect_uri_provided_explicitly=True,
    )
    url = await provider.authorize(client, params)
    state = url.split("state=")[1]
    redirect = provider.complete_entra_callback(code="entra-code", state=state)
    code = redirect.split("code=")[1].split("&")[0]
    loaded = await provider.load_authorization_code(client, code)
    await provider.exchange_authorization_code(client, loaded)
    # Segunda troca falha (consumido).
    from mcp.server.auth.provider import TokenError
    with pytest.raises(TokenError):
        await provider.exchange_authorization_code(client, loaded)


async def test_callback_com_state_desconhecido(provider):
    from mcp_o365.auth.errors import AuthError
    with pytest.raises(AuthError):
        provider.complete_entra_callback(code="x", state="inexistente")
