"""Integração — tools de ESCRITA de calendário end-to-end (US-2.3 a US-2.6).

Cada escrita segue o par prepare/confirm. Invariantes verificadas com o `FakeGraphClient`
a contar chamadas:
- `prepare` NÃO chama o método de escrita (create/update/cancel/respond a 0);
- `confirm` chama o método certo 1× e regista auditoria (event=audit);
- segundo `confirm` é idempotente (não duplica a operação real);
- recorrência sem scope -> needs_clarification (sem token, sem escrita);
- organizador/não-organizador tratados nos casos de cancel/respond (D7);
- reauth graciosa no confirm não chama o Graph nem consome o token.
"""

from __future__ import annotations

import logging
from datetime import timedelta

from mcp_o365.approval.engine import ApprovalEngine
from mcp_o365.auth.plane_b import PlaneB
from mcp_o365.tools import calendar as cal
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


def _approval(store, clock) -> ApprovalEngine:
    return ApprovalEngine(store, clock=clock, ttl_seconds=300)


def _audit_events(caplog) -> list[dict]:
    return [
        getattr(r, "fields", {})
        for r in caplog.records
        if getattr(r, "fields", {}).get("event") == "audit"
    ]


# ============================ US-2.3 — CRIAR ============================


async def test_create_prepare_nao_escreve_confirm_cria_e_audita(
    mapping, store, config, clock, caplog
):
    """US-2.3 — sem location -> Teams; prepare não cria; confirm cria 1× e audita."""
    _link(mapping, clock)
    gc = FakeGraphClient(
        mailbox_timezone="GMT Standard Time",
        event={"id": "evt-novo", "webLink": "https://web/evt-novo"},
    )
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)

    prepared = await cal.run_calendar_create_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, subject_line="Sync", start="2026-06-10T10:00:00",
        end="2026-06-10T11:00:00", attendees=["a@mobiweb.pt", "b@cliente.com"], clock=clock,
    )
    assert prepared["status"] == "pending_confirmation"
    assert prepared["recipients_count"] == 2
    assert "Notifica 2 participante(s)" in prepared["summary"]
    assert "mobiweb.pt" in prepared["summary"] and "cliente.com" in prepared["summary"]
    assert "Inclui link Teams" in prepared["summary"]  # sem location -> Teams (D6)
    # prepare leu o fuso mas NÃO criou o evento.
    assert gc.count("create_event") == 0

    with caplog.at_level(logging.INFO, logger="mcp_o365.audit"):
        confirmed = await cal.run_calendar_create_confirm(
            "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
            approval=approval, confirmation_token=prepared["confirmation_token"],
            clock=clock,
        )
    assert confirmed["status"] == "done"
    assert confirmed["event_id"] == "evt-novo"
    assert confirmed["web_link"] == "https://web/evt-novo"
    assert gc.count("create_event") == 1
    events = _audit_events(caplog)
    audit = next(e for e in events if e["action"] == "calendar.create")
    assert audit["outcome"] == "success"
    assert audit["recipients_count"] == 2
    # Só-metadados: subject_hash presente, nunca o assunto em claro.
    assert audit["subject_hash"] and audit["subject_hash"] != "Sync"
    assert "Sync" not in str(audit)
    assert audit["online"] is True


async def test_create_com_location_e_presencial_sem_teams(mapping, store, config, clock):
    """US-2.3 — com location -> presencial sem link Teams (D6)."""
    _link(mapping, clock)
    gc = FakeGraphClient(mailbox_timezone="GMT Standard Time")
    prepared = await cal.run_calendar_create_prepare(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock), graph_client=gc,
        store=store, approval=_approval(store, clock), subject_line="Almoço",
        start="2026-06-10T12:00:00", end="2026-06-10T13:00:00",
        location="Sala 1", clock=clock,
    )
    assert prepared["status"] == "pending_confirmation"
    assert "Presencial em 'Sala 1' (sem link Teams)" in prepared["summary"]
    assert "Inclui link Teams" not in prepared["summary"]


async def test_create_confirm_idempotente(mapping, store, config, clock):
    """US-2.3 — replay do token não re-cria; create_event fica a 1."""
    _link(mapping, clock)
    gc = FakeGraphClient(mailbox_timezone="GMT Standard Time")
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)
    prepared = await cal.run_calendar_create_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, subject_line="X", start="2026-06-10T10:00:00",
        end="2026-06-10T11:00:00", clock=clock,
    )
    token = prepared["confirmation_token"]
    first = await cal.run_calendar_create_confirm(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, confirmation_token=token, clock=clock,
    )
    second = await cal.run_calendar_create_confirm(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, confirmation_token=token, clock=clock,
    )
    assert first["status"] == "done"
    assert second["idempotent_replay"] is True
    assert gc.count("create_event") == 1


async def test_create_prepare_campos_obrigatorios(mapping, store, config, clock):
    """US-2.3 — falta subject_line/start/end -> error."""
    _link(mapping, clock)
    gc = FakeGraphClient(mailbox_timezone="GMT Standard Time")
    out = await cal.run_calendar_create_prepare(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock), graph_client=gc,
        store=store, approval=_approval(store, clock), subject_line="", start="",
        end="", clock=clock,
    )
    assert out["status"] == "error"
    assert gc.count("create_event") == 0


async def test_create_confirm_reauth_nao_cria_nem_consome_token(
    mapping, store, config, clock
):
    """US-2.3 — refresh falha no confirm -> reauth_required, sem create e token reutilizável."""
    _link(mapping, clock)
    gc = FakeGraphClient(mailbox_timezone="GMT Standard Time")
    approval = _approval(store, clock)
    prepared = await cal.run_calendar_create_prepare(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock), graph_client=gc,
        store=store, approval=approval, subject_line="X", start="2026-06-10T10:00:00",
        end="2026-06-10T11:00:00", clock=clock,
    )
    token = prepared["confirmation_token"]
    # Token expira entre prepare e confirm; o refresh do confirm falha.
    store.update_account_tokens(
        subject="subj-1", account_id="acc-1", access_token="old-at",
        refresh_token="rt-1", expires_at=clock() - timedelta(minutes=5),
    )
    pb_bad = _plane_b(config, clock, refresh_result=INVALID_GRANT_RESPONSE)
    out = await cal.run_calendar_create_confirm(
        "subj-1", mapping=mapping, plane_b=pb_bad, graph_client=gc, store=store,
        approval=approval, confirmation_token=token, clock=clock,
    )
    assert out["status"] == "reauth_required"
    assert gc.count("create_event") == 0
    pending = store.get_pending_operation("subj-1", token)
    assert pending is not None and pending["consumed_at"] is None


# ============================ US-2.4 — EDITAR ============================


async def test_update_nao_recorrente_passa_direto(mapping, store, config, clock, caplog):
    """US-2.4 — evento não-recorrente: prepare normal; confirm faz PATCH e audita."""
    _link(mapping, clock)
    gc = FakeGraphClient(
        mailbox_timezone="GMT Standard Time",
        event={"id": "evt-1", "subject": "Sync", "isRecurring": False,
               "attendees": [{"email": "a@x.com"}]},
    )
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)
    prepared = await cal.run_calendar_update_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, event_id="evt-1", start="2026-06-12T10:00:00", clock=clock,
    )
    assert prepared["status"] == "pending_confirmation"
    assert gc.count("update_event") == 0  # prepare leu mas não escreveu

    with caplog.at_level(logging.INFO, logger="mcp_o365.audit"):
        confirmed = await cal.run_calendar_update_confirm(
            "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
            approval=approval, confirmation_token=prepared["confirmation_token"],
            clock=clock,
        )
    assert confirmed["status"] == "done"
    assert gc.count("update_event") == 1
    # Só os campos alterados entram em changes.
    patch = next(c for c in gc.calls if c[0] == "update_event")
    assert set(patch[2]["changes"].keys()) == {"start"}
    assert patch[1][1] == "evt-1"  # PATCH ao próprio event_id (occurrence/single)
    assert any(e["action"] == "calendar.update" for e in _audit_events(caplog))


async def test_update_recorrente_sem_scope_clarification(mapping, store, config, clock):
    """US-2.4 — recorrente sem scope -> needs_clarification, sem token e sem PATCH."""
    _link(mapping, clock)
    gc = FakeGraphClient(
        mailbox_timezone="GMT Standard Time",
        event={"id": "occ-1", "isRecurring": True, "seriesMasterId": "series-1"},
    )
    out = await cal.run_calendar_update_prepare(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock), graph_client=gc,
        store=store, approval=_approval(store, clock), event_id="occ-1",
        subject_line="Novo", clock=clock,
    )
    assert out["status"] == "needs_clarification"
    assert "confirmation_token" not in out
    assert gc.count("update_event") == 0


async def test_update_recorrente_scope_series_patch_no_seriesmaster(
    mapping, store, config, clock
):
    """US-2.4 — scope='series' -> PATCH ao seriesMasterId; occurrence -> ao próprio id."""
    _link(mapping, clock)
    gc = FakeGraphClient(
        mailbox_timezone="GMT Standard Time",
        event={"id": "occ-1", "isRecurring": True, "seriesMasterId": "series-1",
               "subject": "Daily"},
    )
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)
    prepared = await cal.run_calendar_update_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, event_id="occ-1", subject_line="Daily v2",
        scope="series", clock=clock,
    )
    assert prepared["status"] == "pending_confirmation"
    await cal.run_calendar_update_confirm(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, confirmation_token=prepared["confirmation_token"], clock=clock,
    )
    patch = next(c for c in gc.calls if c[0] == "update_event")
    assert patch[1][1] == "series-1"  # alvo = série


async def test_update_sem_campos_erro(mapping, store, config, clock):
    """US-2.4 — sem campos a alterar -> error."""
    _link(mapping, clock)
    gc = FakeGraphClient(
        mailbox_timezone="GMT Standard Time",
        event={"id": "evt-1", "isRecurring": False},
    )
    out = await cal.run_calendar_update_prepare(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock), graph_client=gc,
        store=store, approval=_approval(store, clock), event_id="evt-1", clock=clock,
    )
    assert out["status"] == "error"


# ============================ US-2.5 — CANCELAR ============================


async def test_cancel_organizador_prepare_confirm_audita(mapping, store, config, clock, caplog):
    """US-2.5 — organizador cancela: resumo declara N participantes; confirm 1× + auditoria."""
    _link(mapping, clock)
    gc = FakeGraphClient(
        mailbox_timezone="GMT Standard Time",
        me={"userPrincipalName": "subj@example.com"},
        event={"id": "evt-1", "subject": "Reunião", "isRecurring": False,
               "organizer": "subj@example.com",
               "attendees": [{"email": "a@x.com"}, {"email": "b@y.com"}]},
    )
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)
    prepared = await cal.run_calendar_cancel_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, event_id="evt-1", message_choice_confirmed=True, clock=clock,
    )
    assert prepared["status"] == "pending_confirmation"
    assert "Notifica 2 participante(s)" in prepared["summary"]
    assert gc.count("cancel_event") == 0  # prepare não cancela

    with caplog.at_level(logging.INFO, logger="mcp_o365.audit"):
        confirmed = await cal.run_calendar_cancel_confirm(
            "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
            approval=approval, confirmation_token=prepared["confirmation_token"],
            clock=clock,
        )
    assert confirmed["status"] == "done"
    assert gc.count("cancel_event") == 1
    assert any(e["action"] == "calendar.cancel" for e in _audit_events(caplog))


async def test_cancel_nao_organizador_orienta_para_decline(mapping, store, config, clock):
    """US-2.5 — não-organizador -> error orientando para decline; sem cancel_event."""
    _link(mapping, clock)
    gc = FakeGraphClient(
        mailbox_timezone="GMT Standard Time",
        me={"userPrincipalName": "subj@example.com"},
        event={"id": "evt-1", "isRecurring": False,
               "organizer": "outro@example.com", "attendees": []},
    )
    out = await cal.run_calendar_cancel_prepare(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock), graph_client=gc,
        store=store, approval=_approval(store, clock), event_id="evt-1", clock=clock,
    )
    assert out["status"] == "error"
    assert "decline" in out["message"]
    assert "confirmation_token" not in out
    assert gc.count("cancel_event") == 0


async def test_cancel_recorrente_sem_scope_clarification(mapping, store, config, clock):
    """US-2.5 — recorrente sem scope -> needs_clarification (sem cancel)."""
    _link(mapping, clock)
    gc = FakeGraphClient(
        mailbox_timezone="GMT Standard Time",
        me={"userPrincipalName": "subj@example.com"},
        event={"id": "occ-1", "isRecurring": True, "seriesMasterId": "series-1",
               "organizer": "subj@example.com", "attendees": []},
    )
    out = await cal.run_calendar_cancel_prepare(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock), graph_client=gc,
        store=store, approval=_approval(store, clock), event_id="occ-1", clock=clock,
    )
    assert out["status"] == "needs_clarification"
    assert gc.count("cancel_event") == 0


async def test_cancel_confirm_idempotente(mapping, store, config, clock):
    """US-2.5 — replay não re-cancela."""
    _link(mapping, clock)
    gc = FakeGraphClient(
        mailbox_timezone="GMT Standard Time",
        me={"userPrincipalName": "subj@example.com"},
        event={"id": "evt-1", "isRecurring": False,
               "organizer": "subj@example.com", "attendees": []},
    )
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)
    prepared = await cal.run_calendar_cancel_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, event_id="evt-1", message_choice_confirmed=True, clock=clock,
    )
    token = prepared["confirmation_token"]
    await cal.run_calendar_cancel_confirm(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, confirmation_token=token, clock=clock,
    )
    second = await cal.run_calendar_cancel_confirm(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, confirmation_token=token, clock=clock,
    )
    assert second["idempotent_replay"] is True
    assert gc.count("cancel_event") == 1


async def test_cancel_sem_escolha_pede_mensagem(mapping, store, config, clock):
    """US-2.5 (melhoria 2026-06-03) — organizador a cancelar sem message_choice_confirmed
    -> needs_clarification (mensagem própria/sugerida/nenhuma); sem token, sem cancel."""
    _link(mapping, clock)
    gc = FakeGraphClient(
        mailbox_timezone="GMT Standard Time",
        me={"userPrincipalName": "subj@example.com"},
        event={"id": "evt-1", "subject": "Reunião Moomenti", "isRecurring": False,
               "organizer": "subj@example.com", "attendees": [{"email": "a@x.com"}]},
    )
    out = await cal.run_calendar_cancel_prepare(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock), graph_client=gc,
        store=store, approval=_approval(store, clock), event_id="evt-1", clock=clock,
    )
    assert out["status"] == "needs_clarification"
    assert "mensagem" in out["question"].lower()
    assert "confirmation_token" not in out
    assert gc.count("cancel_event") == 0
    # 3 opções: minha / sugere / sem mensagem.
    assert len(out["options"]) == 3
    assert any("uge" in o["label"] or "ugest" in o["action"] for o in out["options"])


async def test_cancel_com_mensagem_confirmada(mapping, store, config, clock):
    """US-2.5 — com message_choice_confirmed + comment -> token; confirm cancela com o texto."""
    _link(mapping, clock)
    gc = FakeGraphClient(
        mailbox_timezone="GMT Standard Time",
        me={"userPrincipalName": "subj@example.com"},
        event={"id": "evt-1", "subject": "Reunião Moomenti", "isRecurring": False,
               "organizer": "subj@example.com", "attendees": [{"email": "a@x.com"}]},
    )
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)
    prepared = await cal.run_calendar_cancel_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, event_id="evt-1", comment="Vamos reagendar em breve.",
        message_choice_confirmed=True, clock=clock,
    )
    assert prepared["status"] == "pending_confirmation"
    await cal.run_calendar_cancel_confirm(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, confirmation_token=prepared["confirmation_token"], clock=clock,
    )
    call = next(c for c in gc.calls if c[0] == "cancel_event")
    assert call[2]["comment"] == "Vamos reagendar em breve."


# ============================ US-2.6 — RESPONDER ============================


async def test_respond_declara_transicao_e_confirma(mapping, store, config, clock, caplog):
    """US-2.6 — accept->decline: resumo declara transição; confirm responde 1× + auditoria."""
    _link(mapping, clock)
    gc = FakeGraphClient(
        me={"userPrincipalName": "subj@example.com"},
        event={"id": "evt-1", "subject": "Convite", "organizer": "org@example.com",
               "responseStatus": "accepted"},
    )
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)
    prepared = await cal.run_calendar_respond_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, event_id="evt-1", response="decline",
        comment="não posso", message_choice_confirmed=True, clock=clock,
    )
    assert prepared["status"] == "pending_confirmation"
    assert "Aceitado" in prepared["summary"] and "Recusado" in prepared["summary"]
    assert gc.count("respond_event") == 0

    with caplog.at_level(logging.INFO, logger="mcp_o365.audit"):
        confirmed = await cal.run_calendar_respond_confirm(
            "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
            approval=approval, confirmation_token=prepared["confirmation_token"],
            clock=clock,
        )
    assert confirmed["status"] == "done"
    assert gc.count("respond_event") == 1
    call = next(c for c in gc.calls if c[0] == "respond_event")
    assert call[2]["response"] == "decline"
    audit = next(e for e in _audit_events(caplog) if e["action"] == "calendar.respond")
    assert audit["response"] == "decline"
    assert audit["previous"] == "accepted"


async def test_respond_organizador_bloqueado(mapping, store, config, clock):
    """US-2.6 — o organizador não responde ao próprio convite -> error (sem token)."""
    _link(mapping, clock)
    gc = FakeGraphClient(
        me={"userPrincipalName": "subj@example.com"},
        event={"id": "evt-1", "organizer": "subj@example.com",
               "responseStatus": "organizer"},
    )
    out = await cal.run_calendar_respond_prepare(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock), graph_client=gc,
        store=store, approval=_approval(store, clock), event_id="evt-1",
        response="accept", clock=clock,
    )
    assert out["status"] == "error"
    assert "confirmation_token" not in out
    assert gc.count("respond_event") == 0


async def test_respond_invalida_erro(mapping, store, config, clock):
    """US-2.6 — response inválida -> error (sem ler o evento)."""
    _link(mapping, clock)
    gc = FakeGraphClient(me={"userPrincipalName": "subj@example.com"})
    out = await cal.run_calendar_respond_prepare(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock), graph_client=gc,
        store=store, approval=_approval(store, clock), event_id="evt-1",
        response="talvez-nao-sei", clock=clock,
    )
    assert out["status"] == "error"
    assert gc.count("respond_event") == 0


async def test_respond_confirm_idempotente(mapping, store, config, clock):
    """US-2.6 — replay não re-responde."""
    _link(mapping, clock)
    gc = FakeGraphClient(
        me={"userPrincipalName": "subj@example.com"},
        event={"id": "evt-1", "organizer": "org@example.com",
               "responseStatus": "none"},
    )
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)
    prepared = await cal.run_calendar_respond_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, event_id="evt-1", response="accept", clock=clock,
    )
    token = prepared["confirmation_token"]
    await cal.run_calendar_respond_confirm(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, confirmation_token=token, clock=clock,
    )
    second = await cal.run_calendar_respond_confirm(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, confirmation_token=token, clock=clock,
    )
    assert second["idempotent_replay"] is True
    assert gc.count("respond_event") == 1


# ===================== TRANSVERSAIS =====================


async def test_token_expirado_devolve_expired(mapping, store, config, clock):
    """Transversal — TTL expirado no confirm -> expired."""
    _link(mapping, clock)
    gc = FakeGraphClient(mailbox_timezone="GMT Standard Time")
    approval = ApprovalEngine(store, clock=clock, ttl_seconds=60)
    pb = _plane_b(config, clock)
    prepared = await cal.run_calendar_create_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, subject_line="X", start="2026-06-10T10:00:00",
        end="2026-06-10T11:00:00", clock=clock,
    )
    clock.advance(120)  # passa o TTL
    out = await cal.run_calendar_create_confirm(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, confirmation_token=prepared["confirmation_token"], clock=clock,
    )
    assert out["status"] == "expired"
    assert gc.count("create_event") == 0


async def test_token_de_outro_subject_e_rejeitado(mapping, store, config, clock):
    """Transversal — token de outro subject -> error (ConfirmationNotFound)."""
    _link(mapping, clock)
    mapping.link_account(
        subject="subj-2", access_token="valid-at", refresh_token="rt-1",
        expires_at=clock() + timedelta(hours=1), home_account_id="acc-2",
    )
    gc = FakeGraphClient(mailbox_timezone="GMT Standard Time")
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)
    prepared = await cal.run_calendar_create_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, subject_line="X", start="2026-06-10T10:00:00",
        end="2026-06-10T11:00:00", clock=clock,
    )
    out = await cal.run_calendar_create_confirm(
        "subj-2", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, confirmation_token=prepared["confirmation_token"], clock=clock,
    )
    assert out["status"] == "error"
    assert gc.count("create_event") == 0


# --- Melhoria 2026-06-03: recusar pergunta pela mensagem ---


async def test_respond_decline_sem_escolha_pede_mensagem(mapping, store, config, clock):
    """US-2.6 — decline sem message_choice_confirmed -> needs_clarification (sem token,
    sem responder). Devolve as opções de mensagem."""
    _link(mapping, clock)
    gc = FakeGraphClient(
        me={"userPrincipalName": "subj@example.com"},
        event={"id": "evt-1", "subject": "Convite", "organizer": "org@example.com",
               "responseStatus": "none"},
    )
    out = await cal.run_calendar_respond_prepare(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock), graph_client=gc,
        store=store, approval=_approval(store, clock), event_id="evt-1",
        response="decline", clock=clock,
    )
    assert out["status"] == "needs_clarification"
    assert "mensagem" in out["question"].lower()
    assert "confirmation_token" not in out
    assert gc.count("respond_event") == 0
    # Tem as 3 opções (com mensagem, sem mensagem, sem notificar).
    assert len(out["options"]) == 3


async def test_respond_decline_com_mensagem(mapping, store, config, clock):
    """US-2.6 — decline confirmado com texto -> respond_event com comment e send_response=True."""
    _link(mapping, clock)
    gc = FakeGraphClient(
        me={"userPrincipalName": "subj@example.com"},
        event={"id": "evt-1", "subject": "Convite", "organizer": "org@example.com",
               "responseStatus": "none"},
    )
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)
    prepared = await cal.run_calendar_respond_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, event_id="evt-1", response="decline",
        comment="Não consigo participar, obrigado.", message_choice_confirmed=True,
        clock=clock,
    )
    assert prepared["status"] == "pending_confirmation"
    assert "mensagem" in prepared["summary"].lower()
    await cal.run_calendar_respond_confirm(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, confirmation_token=prepared["confirmation_token"], clock=clock,
    )
    call = next(c for c in gc.calls if c[0] == "respond_event")
    assert call[2]["comment"] == "Não consigo participar, obrigado."
    assert call[2]["send_response"] is True


async def test_respond_decline_sem_notificar(mapping, store, config, clock):
    """US-2.6 — decline confirmado com notify_organizer=false -> send_response=False."""
    _link(mapping, clock)
    gc = FakeGraphClient(
        me={"userPrincipalName": "subj@example.com"},
        event={"id": "evt-1", "subject": "Convite", "organizer": "org@example.com",
               "responseStatus": "none"},
    )
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)
    prepared = await cal.run_calendar_respond_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, event_id="evt-1", response="decline",
        message_choice_confirmed=True, notify_organizer=False, clock=clock,
    )
    assert prepared["status"] == "pending_confirmation"
    assert "sem notificar" in prepared["summary"].lower()
    await cal.run_calendar_respond_confirm(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, confirmation_token=prepared["confirmation_token"], clock=clock,
    )
    call = next(c for c in gc.calls if c[0] == "respond_event")
    assert call[2]["send_response"] is False


async def test_respond_accept_nao_pergunta_mensagem(mapping, store, config, clock):
    """US-2.6 — accept NÃO dispara a pergunta da mensagem (vai direto a pending_confirmation)."""
    _link(mapping, clock)
    gc = FakeGraphClient(
        me={"userPrincipalName": "subj@example.com"},
        event={"id": "evt-1", "organizer": "org@example.com", "responseStatus": "none"},
    )
    out = await cal.run_calendar_respond_prepare(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock), graph_client=gc,
        store=store, approval=_approval(store, clock), event_id="evt-1",
        response="accept", clock=clock,
    )
    assert out["status"] == "pending_confirmation"
