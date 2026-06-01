"""Integração — whoami end-to-end (T11) com Graph (respx) e Entra (FakeMsalApp) mockados.

Cobre os três caminhos críticos: sessão válida, refresh-no-meio, e invalid_grant ->
reautenticação graciosa (o comportamento exigido pelo bloqueador de Conditional Access).
"""

from __future__ import annotations

from datetime import timedelta

import httpx
import respx

from mcp_o365.auth.plane_b import PlaneB
from mcp_o365.graph.client import GraphClient
from mcp_o365.tools.whoami import run_whoami
from tests.conftest import INVALID_GRANT_RESPONSE, FakeMsalApp, graph_token_response

GRAPH_ME = "https://graph.microsoft.com/v1.0/me"
ME_BODY = {"id": "graph-id-1", "displayName": "Márcio", "userPrincipalName": "marcio@example.com"}


async def _noop_sleep(_seconds: float) -> None:
    return None


def _graph_client() -> GraphClient:
    return GraphClient(httpx.AsyncClient(), sleeper=_noop_sleep)


def _plane_b(config, clock, refresh_result=None) -> PlaneB:
    result = refresh_result or graph_token_response(refresh_token="rt-new")
    fake = FakeMsalApp(refresh_result=result)
    return PlaneB(config, msal_app_factory=lambda: fake, clock=clock)


@respx.mock
async def test_sessao_valida_sem_refresh(mapping, store, config, clock):
    respx.get(GRAPH_ME).mock(return_value=httpx.Response(200, json=ME_BODY))
    mapping.link_account(
        subject="subj-1", access_token="valid-at", refresh_token="rt-1",
        expires_at=clock() + timedelta(hours=1), home_account_id="acc-1",
    )
    out = await run_whoami(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock),
        graph_client=_graph_client(), store=store, clock=clock,
    )
    assert out["status"] == "ok"
    assert out["userPrincipalName"] == "marcio@example.com"


@respx.mock
async def test_refresh_no_meio(mapping, store, config, clock):
    respx.get(GRAPH_ME).mock(return_value=httpx.Response(200, json=ME_BODY))
    # Token já expirado -> tem de refrescar antes de chamar o Graph.
    mapping.link_account(
        subject="subj-1", access_token="old-at", refresh_token="rt-1",
        expires_at=clock() - timedelta(minutes=5), home_account_id="acc-1",
    )
    out = await run_whoami(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock),
        graph_client=_graph_client(), store=store, clock=clock,
    )
    assert out["status"] == "ok"
    # O token novo foi persistido.
    acc = store.get_account("subj-1", "acc-1")
    assert acc["access_token"] == "graph-access-1"
    assert acc["refresh_token"] == "rt-new"


@respx.mock
async def test_invalid_grant_reauth_graciosa(mapping, store, config, clock):
    # O Graph nem chega a ser chamado: o refresh falha primeiro.
    mapping.link_account(
        subject="subj-1", access_token="old-at", refresh_token="rt-1",
        expires_at=clock() - timedelta(minutes=5), home_account_id="acc-1",
    )
    pb = _plane_b(config, clock, refresh_result=INVALID_GRANT_RESPONSE)
    out = await run_whoami(
        "subj-1", mapping=mapping, plane_b=pb,
        graph_client=_graph_client(), store=store, clock=clock,
    )
    assert out["status"] == "reauth_required"
    # A sessão foi marcada como expirada (força novo login).
    assert mapping.select_account("subj-1") is None


async def test_sem_subject_pede_reauth(mapping, store, config, clock):
    out = await run_whoami(
        None, mapping=mapping, plane_b=_plane_b(config, clock),
        graph_client=_graph_client(), store=store, clock=clock,
    )
    assert out["status"] == "reauth_required"


async def test_sem_conta_ligada_pede_reauth(mapping, store, config, clock):
    out = await run_whoami(
        "subj-sem-conta", mapping=mapping, plane_b=_plane_b(config, clock),
        graph_client=_graph_client(), store=store, clock=clock,
    )
    assert out["status"] == "reauth_required"
