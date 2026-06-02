"""Unit — resolve_access_token (_session.py): os 4 casos de ReauthRequired, caso feliz e
refresh proativo persistido.

Tudo mockado: Entra via FakeMsalApp, sem rede nem sleep.
"""

from __future__ import annotations

from datetime import timedelta

import pytest

from mcp_o365.auth.errors import ReauthRequired
from mcp_o365.auth.plane_b import PlaneB
from mcp_o365.tools._session import resolve_access_token
from tests.conftest import INVALID_GRANT_RESPONSE, FakeMsalApp, graph_token_response


def _plane_b(config, clock, refresh_result=None) -> PlaneB:
    result = refresh_result or graph_token_response(refresh_token="rt-new")
    fake = FakeMsalApp(refresh_result=result)
    return PlaneB(config, msal_app_factory=lambda: fake, clock=clock)


async def test_reauth_sem_subject(mapping, store, config, clock):
    with pytest.raises(ReauthRequired):
        await resolve_access_token(
            None, mapping=mapping, plane_b=_plane_b(config, clock), store=store, clock=clock
        )


async def test_reauth_sem_conta_ligada(mapping, store, config, clock):
    with pytest.raises(ReauthRequired):
        await resolve_access_token(
            "subj-sem-conta", mapping=mapping, plane_b=_plane_b(config, clock),
            store=store, clock=clock,
        )


async def test_reauth_sem_refresh_token(mapping, store, config, clock):
    """Token expirado + sem refresh token -> reauth (e conta marcada expirada)."""
    mapping.link_account(
        subject="subj-1", access_token="old-at", refresh_token=None,
        expires_at=clock() - timedelta(minutes=5), home_account_id="acc-1",
    )
    with pytest.raises(ReauthRequired):
        await resolve_access_token(
            "subj-1", mapping=mapping, plane_b=_plane_b(config, clock),
            store=store, clock=clock,
        )
    assert mapping.select_account("subj-1") is None


async def test_reauth_refresh_invalid_grant(mapping, store, config, clock):
    """Refresh rejeitado (invalid_grant pela CA) -> reauth graciosa."""
    mapping.link_account(
        subject="subj-1", access_token="old-at", refresh_token="rt-1",
        expires_at=clock() - timedelta(minutes=5), home_account_id="acc-1",
    )
    pb = _plane_b(config, clock, refresh_result=INVALID_GRANT_RESPONSE)
    with pytest.raises(ReauthRequired):
        await resolve_access_token(
            "subj-1", mapping=mapping, plane_b=pb, store=store, clock=clock
        )
    assert mapping.select_account("subj-1") is None


async def test_caso_feliz_devolve_conta_e_token(mapping, store, config, clock):
    mapping.link_account(
        subject="subj-1", access_token="valid-at", refresh_token="rt-1",
        expires_at=clock() + timedelta(hours=1), home_account_id="acc-1",
    )
    account, token = await resolve_access_token(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock),
        store=store, clock=clock,
    )
    assert account.account_id == "acc-1"
    assert token == "valid-at"  # sem refresh: devolve o token corrente


async def test_refresh_proativo_persiste_novo_token(mapping, store, config, clock):
    """Expira dentro de 60s -> renova proativamente e persiste o token novo no store."""
    mapping.link_account(
        subject="subj-1", access_token="old-at", refresh_token="rt-1",
        expires_at=clock() + timedelta(seconds=30), home_account_id="acc-1",
    )
    _account, token = await resolve_access_token(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock),
        store=store, clock=clock,
    )
    assert token == "graph-access-1"
    acc = store.get_account("subj-1", "acc-1")
    assert acc["access_token"] == "graph-access-1"
    assert acc["refresh_token"] == "rt-new"


# ---- call_graph: resiliência a 401/403 do Graph (refresh forçado + retry) ----

from mcp_o365.auth.errors import UpstreamAuthError  # noqa: E402
from mcp_o365.tools._session import call_graph  # noqa: E402


async def test_call_graph_401_forca_refresh_e_repete_com_sucesso(mapping, store, config, clock):
    """Token aparentemente válido recusado pelo Graph (401) -> refresh forçado + retry OK."""
    mapping.link_account(
        subject="subj-1", access_token="valid-at", refresh_token="rt-1",
        expires_at=clock() + timedelta(hours=1), home_account_id="acc-1",
    )
    tokens_vistos: list[str] = []

    async def op(token: str) -> str:
        tokens_vistos.append(token)
        if token == "valid-at":
            raise UpstreamAuthError("Graph rejeitou o token (401).")
        return "ok"

    account, result = await call_graph(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock),
        store=store, op=op, clock=clock,
    )
    assert result == "ok"
    assert account.account_id == "acc-1"
    # 1ª tentativa com o token corrente (falha), 2ª com o token renovado (sucesso).
    assert tokens_vistos == ["valid-at", "graph-access-1"]
    assert store.get_account("subj-1", "acc-1")["access_token"] == "graph-access-1"


async def test_call_graph_401_persistente_vira_reauth(mapping, store, config, clock):
    """Mesmo após refresh, o Graph continua a recusar -> ReauthRequired (graciosa)."""
    mapping.link_account(
        subject="subj-1", access_token="valid-at", refresh_token="rt-1",
        expires_at=clock() + timedelta(hours=1), home_account_id="acc-1",
    )

    async def op(_token: str):
        raise UpstreamAuthError("Graph rejeitou o token (403).")

    with pytest.raises(ReauthRequired):
        await call_graph(
            "subj-1", mapping=mapping, plane_b=_plane_b(config, clock),
            store=store, op=op, clock=clock,
        )
    assert mapping.select_account("subj-1") is None  # conta marcada expirada
