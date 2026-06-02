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
