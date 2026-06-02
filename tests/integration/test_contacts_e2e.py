"""Integração — resolução de destinatários por nome (US-5.x).

Verifica: 0 -> not_found; 1 -> ok com recipient; vários -> needs_clarification; dedupe
People+Contactos por email; reauth graciosa quando o refresh falha.
"""

from __future__ import annotations

from datetime import timedelta

from mcp_o365.auth.plane_b import PlaneB
from mcp_o365.tools.contacts import run_resolve_recipient
from tests.conftest import INVALID_GRANT_RESPONSE, FakeMsalApp, graph_token_response
from tests.integration.fake_graph import FakeGraphClient


def _plane_b(config, clock, refresh_result=None) -> PlaneB:
    result = refresh_result or graph_token_response(refresh_token="rt-new")
    fake = FakeMsalApp(refresh_result=result)
    return PlaneB(config, msal_app_factory=lambda: fake, clock=clock)


def _link(mapping, clock) -> None:
    mapping.link_account(
        subject="subj-1", access_token="valid-at", refresh_token="rt-1",
        expires_at=clock() + timedelta(hours=1), home_account_id="acc-1",
    )


def _person(name, email):
    return {"display_name": name, "email": email, "source": "people"}


def _contact(name, email):
    return {"display_name": name, "email": email, "source": "contacts"}


async def test_nome_sem_resultados(mapping, store, config, clock):
    _link(mapping, clock)
    gc = FakeGraphClient(people=[], contacts=[])
    out = await run_resolve_recipient(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock), graph_client=gc,
        store=store, name="ninguem", clock=clock,
    )
    assert out["status"] == "not_found"
    assert out["candidates"] == []


async def test_um_unico_candidato(mapping, store, config, clock):
    _link(mapping, clock)
    gc = FakeGraphClient(people=[_person("Vera Costa", "vera.costa@mobiweb.pt")])
    out = await run_resolve_recipient(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock), graph_client=gc,
        store=store, name="vera", clock=clock,
    )
    assert out["status"] == "ok"
    assert out["recipient"]["email"] == "vera.costa@mobiweb.pt"


async def test_varios_candidatos_pede_clarificacao(mapping, store, config, clock):
    _link(mapping, clock)
    gc = FakeGraphClient(
        people=[_person("Vera Costa", "vera.costa@mobiweb.pt")],
        contacts=[_contact("Vera Nunes", "vera.nunes@habisonho.com")],
    )
    out = await run_resolve_recipient(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock), graph_client=gc,
        store=store, name="vera", clock=clock,
    )
    assert out["status"] == "needs_clarification"
    assert len(out["candidates"]) == 2
    emails = {c["email"] for c in out["candidates"]}
    assert emails == {"vera.costa@mobiweb.pt", "vera.nunes@habisonho.com"}


async def test_dedupe_people_e_contactos_pelo_email(mapping, store, config, clock):
    _link(mapping, clock)
    # Mesma pessoa em People e Contactos (email igual, caixa diferente) -> 1 candidato.
    gc = FakeGraphClient(
        people=[_person("Vera Costa", "Vera.Costa@mobiweb.pt")],
        contacts=[_contact("Vera", "vera.costa@mobiweb.pt")],
    )
    out = await run_resolve_recipient(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock), graph_client=gc,
        store=store, name="vera", clock=clock,
    )
    assert out["status"] == "ok"  # dedupe -> 1
    assert out["recipient"]["source"] == "people"  # People tem prioridade


async def test_candidatos_sem_email_sao_ignorados(mapping, store, config, clock):
    _link(mapping, clock)
    gc = FakeGraphClient(people=[_person("Sem Email", None), _person("Com Email", "x@y.pt")])
    out = await run_resolve_recipient(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock), graph_client=gc,
        store=store, name="x", clock=clock,
    )
    assert out["status"] == "ok"
    assert out["recipient"]["email"] == "x@y.pt"


async def test_reauth_quando_refresh_falha(mapping, store, config, clock):
    mapping.link_account(
        subject="subj-1", access_token="old", refresh_token="rt-1",
        expires_at=clock() - timedelta(minutes=5), home_account_id="acc-1",
    )
    pb = _plane_b(config, clock, refresh_result=INVALID_GRANT_RESPONSE)
    gc = FakeGraphClient(people=[_person("Vera", "v@x.pt")])
    out = await run_resolve_recipient(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc,
        store=store, name="vera", clock=clock,
    )
    assert out["status"] == "reauth_required"
    assert gc.count("search_people") == 0  # nem chegou a chamar o Graph
