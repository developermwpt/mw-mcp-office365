"""Integração — fluxo de aprendizagem ponta-a-ponta (US-L.x).

Fluxo verificado:
1. opt-in;
2. registar comportamentos via os `*_confirm` de email (com `FakeGraphClient`);
3. `email_recommendations` devolve uma sugestão coerente;
4. aceitar a sugestão -> chamar o `*_prepare` indicado -> `confirmation_token`;
5. `*_confirm` executa (reutilizando o `fake_graph`).

Verifica também a invariante de privacidade: SEM opt-in não há registo nem recomendações,
e o registo de comportamento NUNCA quebra a operação de email.
"""

from __future__ import annotations

from datetime import timedelta

from mcp_o365.approval.engine import ApprovalEngine
from mcp_o365.auth.plane_b import PlaneB
from mcp_o365.learning.recommender import Recommender
from mcp_o365.tools.email import (
    run_email_move_confirm,
    run_email_move_prepare,
)
from mcp_o365.tools.learning import (
    run_email_recommendations,
    run_learning_forget,
    run_learning_opt_in,
)
from tests.conftest import FakeMsalApp, graph_token_response
from tests.integration.fake_graph import FakeGraphClient


def _plane_b(config, clock) -> PlaneB:
    fake = FakeMsalApp(refresh_result=graph_token_response(refresh_token="rt-new"))
    return PlaneB(config, msal_app_factory=lambda: fake, clock=clock)


def _link(mapping, clock) -> None:
    mapping.link_account(
        subject="subj-1", access_token="valid-at", refresh_token="rt-1",
        expires_at=clock() + timedelta(hours=1), home_account_id="acc-1",
    )


def _approval(store, clock) -> ApprovalEngine:
    return ApprovalEngine(store, clock=clock, ttl_seconds=300)


def _meta(domain: str = "newsletter.acme.com") -> dict:
    return {
        "from": {"emailAddress": {"address": f"news@{domain}"}},
        "subject": "Promoção semanal",
        "hasAttachments": False,
        "internetMessageHeaders": [{"name": "List-Id", "value": "acme"}],
    }


async def _move_archive(mapping, store, config, clock, *, message_id: str) -> None:
    """Move (arquiva) um email via prepare/confirm, com metadados para a aprendizagem."""
    gc = FakeGraphClient()
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)
    prepared = await run_email_move_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, message_id=message_id, destination="Archive",
        message_meta=_meta(), clock=clock,
    )
    await run_email_move_confirm(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, confirmation_token=prepared["confirmation_token"], clock=clock,
    )


async def test_fluxo_opt_in_aprende_recomenda_e_executa(mapping, store, config, clock):
    _link(mapping, clock)

    # 1) opt-in
    opt = await run_learning_opt_in("subj-1", store=store, enabled=True, clock=clock)
    assert opt["status"] == "ok" and opt["opt_in"] is True

    # 2) registar comportamento: arquivar várias vezes emails do mesmo domínio
    for i in range(4):
        await _move_archive(mapping, store, config, clock, message_id=f"m{i}")
    assert len(store.list_behavior_events("subj-1")) == 4

    # 3) recomendações para um NOVO email do mesmo domínio
    recommender = Recommender(min_confidence=0.4, top_n=3)
    out = await run_email_recommendations(
        "subj-1", store=store, recommender=recommender,
        message=_meta(), message_id="m-novo", clock=clock,
    )
    assert out["status"] == "ok"
    assert out["count"] >= 1
    top = out["recommendations"][0]
    assert top["action"] == "archive"
    assert "newsletter.acme.com" in top["rationale"]

    # 4) aceitar -> chamar o prepare_tool indicado com os prepare_params
    assert top["prepare_tool"] == "email_move_prepare"
    gc = FakeGraphClient()
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)
    prepared = await run_email_move_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, message_id="m-novo",
        destination=top["prepare_params"]["destination"], clock=clock,
    )
    assert prepared["status"] == "pending_confirmation"
    assert prepared["confirmation_token"]

    # 5) confirmar executa o move de verdade (1x no fake graph)
    done = await run_email_move_confirm(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, confirmation_token=prepared["confirmation_token"], clock=clock,
    )
    assert done["status"] == "done"
    assert gc.count("move_message") == 1


async def test_sem_opt_in_nao_regista_nem_recomenda(mapping, store, config, clock):
    _link(mapping, clock)
    # SEM opt-in: mover não regista comportamento.
    for i in range(4):
        await _move_archive(mapping, store, config, clock, message_id=f"m{i}")
    assert store.list_behavior_events("subj-1") == []

    # E recomendações devolvem opt_out com mensagem de ativação.
    recommender = Recommender()
    out = await run_email_recommendations(
        "subj-1", store=store, recommender=recommender, message=_meta(), clock=clock,
    )
    assert out["status"] == "opt_out"
    assert out["recommendations"] == []
    assert "learning_opt_in" in out["message"]


async def test_forget_apaga_historico(mapping, store, config, clock):
    _link(mapping, clock)
    await run_learning_opt_in("subj-1", store=store, enabled=True, clock=clock)
    for i in range(3):
        await _move_archive(mapping, store, config, clock, message_id=f"m{i}")
    assert len(store.list_behavior_events("subj-1")) == 3

    forget = await run_learning_forget("subj-1", store=store, clock=clock)
    assert forget["status"] == "ok"
    assert forget["deleted"] == 3
    assert store.list_behavior_events("subj-1") == []


async def test_registo_de_comportamento_nunca_quebra_o_email(mapping, store, config, clock):
    """Se o store falhar a registar, o move principal continua a concluir."""
    _link(mapping, clock)
    await run_learning_opt_in("subj-1", store=store, enabled=True, clock=clock)

    # Sabota o registo: record_behavior_event passa a levantar.
    def _boom(*args, **kwargs):
        raise RuntimeError("disco cheio (simulado)")

    store.record_behavior_event = _boom  # type: ignore[method-assign]

    gc = FakeGraphClient()
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)
    prepared = await run_email_move_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, message_id="m1", destination="Archive",
        message_meta=_meta(), clock=clock,
    )
    done = await run_email_move_confirm(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, confirmation_token=prepared["confirmation_token"], clock=clock,
    )
    # O move concluiu apesar da falha de registo (degradação graciosa).
    assert done["status"] == "done"
    assert gc.count("move_message") == 1


async def test_dismiss_suprime_recomendacao(mapping, store, config, clock):
    from mcp_o365.tools.learning import run_learning_dismiss
    _link(mapping, clock)
    await run_learning_opt_in("subj-1", store=store, enabled=True, clock=clock)
    for i in range(4):
        await _move_archive(mapping, store, config, clock, message_id=f"m{i}")

    await run_learning_dismiss(
        "subj-1", store=store, action="archive",
        sender_domain="newsletter.acme.com", clock=clock,
    )
    recommender = Recommender(min_confidence=0.4, top_n=3)
    out = await run_email_recommendations(
        "subj-1", store=store, recommender=recommender,
        message=_meta(), message_id="m-novo", clock=clock,
    )
    assert all(r["action"] != "archive" for r in out["recommendations"])


async def test_purge_expired_remove_antigos(mapping, store, config, clock):
    from mcp_o365.tools.learning import run_learning_purge_expired
    _link(mapping, clock)
    await run_learning_opt_in("subj-1", store=store, enabled=True, clock=clock)
    for i in range(2):
        await _move_archive(mapping, store, config, clock, message_id=f"m{i}")
    assert len(store.list_behavior_events("subj-1")) == 2

    clock.advance(200 * 86400)  # 200 dias depois
    out = await run_learning_purge_expired(
        "subj-1", store=store, retention_days=180, clock=clock
    )
    assert out["status"] == "ok" and out["deleted"] == 2
    assert store.list_behavior_events("subj-1") == []
