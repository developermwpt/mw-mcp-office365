"""Integração — tools de ESCRITA de email end-to-end (US-1.3/1.4/1.6/1.7/1.8).

Cada escrita segue o par prepare/confirm. Invariantes verificadas com o `FakeGraphClient`
a contar chamadas:
- `prepare` devolve token e NÃO toca no Graph;
- `confirm` chama o método Graph certo exatamente 1x e regista auditoria (event=audit);
- segundo `confirm` é idempotente (não duplica a operação real).
"""

from __future__ import annotations

import logging
from datetime import timedelta

import pytest

from mcp_o365.approval.engine import ApprovalEngine
from mcp_o365.auth.plane_b import PlaneB
from mcp_o365.tools.email import (
    _DEFERRED_SEND_PROP_ID,
    run_email_delete_confirm,
    run_email_delete_prepare,
    run_email_move_confirm,
    run_email_move_prepare,
    run_email_reply_confirm,
    run_email_reply_prepare,
    run_email_schedule_cancel_confirm,
    run_email_schedule_cancel_prepare,
    run_email_schedule_confirm,
    run_email_schedule_prepare,
    run_email_send_confirm,
    run_email_send_prepare,
)
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


# ============================ US-1.3 — ENVIAR ============================


async def test_send_prepare_nao_toca_graph_confirm_envia_e_audita(
    mapping, store, config, clock, caplog
):
    _link(mapping, clock)
    gc = FakeGraphClient()
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)

    prepared = await run_email_send_prepare(
        "subj-1", mapping=mapping, plane_b=pb, store=store, approval=approval,
        to=["dest@example.com"], body="Olá", subject_line="Teste", clock=clock,
    )
    assert prepared["status"] == "pending_confirmation"
    assert prepared["confirmation_token"]
    assert prepared["recipients_count"] == 1
    assert prepared["large_attachments"] is False
    # prepare NÃO chamou nada no Graph.
    assert gc.calls == []

    with caplog.at_level(logging.INFO, logger="mcp_o365.audit"):
        confirmed = await run_email_send_confirm(
            "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
            approval=approval, confirmation_token=prepared["confirmation_token"],
            clock=clock,
        )
    assert confirmed["status"] == "done"
    assert gc.count("send_mail") == 1
    events = _audit_events(caplog)
    assert any(e["action"] == "email.send" and e["outcome"] == "success" for e in events)


async def test_send_confirm_idempotente_nao_duplica(mapping, store, config, clock):
    _link(mapping, clock)
    gc = FakeGraphClient()
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)

    prepared = await run_email_send_prepare(
        "subj-1", mapping=mapping, plane_b=pb, store=store, approval=approval,
        to=["dest@example.com"], body="Olá", clock=clock,
    )
    token = prepared["confirmation_token"]
    first = await run_email_send_confirm(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, confirmation_token=token, clock=clock,
    )
    second = await run_email_send_confirm(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, confirmation_token=token, clock=clock,
    )
    assert first["status"] == "done"
    assert second["status"] == "done"
    assert second["idempotent_replay"] is True
    # O envio real só aconteceu UMA vez.
    assert gc.count("send_mail") == 1


async def test_send_prepare_sem_to_erro(mapping, store, config, clock):
    _link(mapping, clock)
    out = await run_email_send_prepare(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock), store=store,
        approval=_approval(store, clock), to=[], body="x", clock=clock,
    )
    assert out["status"] == "error"


# ===================== US-1.6 — ANEXO GRANDE (>3MB) =====================


async def test_send_anexo_grande_marca_flag_e_segue_caminho_draft(
    mapping, store, config, clock
):
    _link(mapping, clock)
    gc = FakeGraphClient(draft={"id": "draft-9"})
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)

    big = {
        "name": "grande.bin",
        "contentBytes": "QUJD",
        "size": 4 * 1024 * 1024,  # 4 MB > limite inline de 3 MB
    }
    prepared = await run_email_send_prepare(
        "subj-1", mapping=mapping, plane_b=pb, store=store, approval=approval,
        to=["dest@example.com"], body="Olá", attachments=[big], clock=clock,
    )
    assert prepared["large_attachments"] is True

    confirmed = await run_email_send_confirm(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, confirmation_token=prepared["confirmation_token"], clock=clock,
    )
    assert confirmed["status"] == "done"
    # Caminho de draft: cria rascunho + abre upload session + envia rascunho.
    # NÃO usa o sendMail inline.
    assert gc.count("create_draft") == 1
    assert gc.count("create_attachment_upload_session") == 1
    assert gc.count("upload_attachment_bytes") == 1
    assert gc.count("send_draft") == 1
    assert gc.count("send_mail") == 0
    # Os bytes carregados são os do anexo, descodificados de base64 ("QUJD" -> b"ABC").
    upload = next(c for c in gc.calls if c[0] == "upload_attachment_bytes")
    assert upload[2]["content_bytes"] == b"ABC"


# ===================== US-1.4 — RESPONDER / REENCAMINHAR =====================


async def test_reply_prepare_confirm_e_idempotente(mapping, store, config, clock):
    _link(mapping, clock)
    gc = FakeGraphClient()
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)

    prepared = await run_email_reply_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, message_id="m1", comment="Obrigado", mode="reply", clock=clock,
    )
    assert prepared["status"] == "pending_confirmation"
    # prepare pode LER (para contar destinatários) mas nunca ESCREVE antes do confirm.
    assert gc.count("reply") == 0 and gc.count("forward") == 0

    token = prepared["confirmation_token"]
    await run_email_reply_confirm(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, confirmation_token=token, clock=clock,
    )
    await run_email_reply_confirm(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, confirmation_token=token, clock=clock,
    )
    assert gc.count("reply") == 1  # idempotente


async def test_reply_all(mapping, store, config, clock):
    _link(mapping, clock)
    gc = FakeGraphClient()
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)
    prepared = await run_email_reply_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, message_id="m1", comment="Para todos", mode="reply_all",
        clock=clock,
    )
    await run_email_reply_confirm(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, confirmation_token=prepared["confirmation_token"], clock=clock,
    )
    assert gc.count("reply") == 1
    # reply_all -> reply com reply_all=True
    reply_call = next(c for c in gc.calls if c[0] == "reply")
    assert reply_call[2]["reply_all"] is True


async def test_forward(mapping, store, config, clock):
    _link(mapping, clock)
    gc = FakeGraphClient()
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)
    prepared = await run_email_reply_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, message_id="m1", comment="FYI", mode="forward",
        to_recipients=["novo@example.com"], clock=clock,
    )
    assert prepared["status"] == "pending_confirmation"
    await run_email_reply_confirm(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, confirmation_token=prepared["confirmation_token"], clock=clock,
    )
    assert gc.count("forward") == 1
    fwd = next(c for c in gc.calls if c[0] == "forward")
    assert fwd[2]["to_recipients"] == ["novo@example.com"]


async def test_forward_sem_destinatarios_erro(mapping, store, config, clock):
    _link(mapping, clock)
    out = await run_email_reply_prepare(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock),
        graph_client=FakeGraphClient(), store=store,
        approval=_approval(store, clock), message_id="m1", comment="x",
        mode="forward", clock=clock,
    )
    assert out["status"] == "error"


async def test_reply_mode_invalido_erro(mapping, store, config, clock):
    _link(mapping, clock)
    out = await run_email_reply_prepare(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock),
        graph_client=FakeGraphClient(), store=store,
        approval=_approval(store, clock), message_id="m1", comment="x",
        mode="modo-que-nao-existe", clock=clock,
    )
    assert out["status"] == "error"


# ============================ US-1.7 — MOVER ============================


async def test_move_prepare_confirm_e_idempotente(mapping, store, config, clock, caplog):
    _link(mapping, clock)
    gc = FakeGraphClient(moved={"id": "m1-na-archive"})
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)

    prepared = await run_email_move_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, message_id="m1", destination="Archive", clock=clock,
    )
    assert prepared["status"] == "pending_confirmation"
    # prepare pode resolver a pasta, mas a pasta "archive" é bem-conhecida -> sem move_message.
    assert gc.count("move_message") == 0

    token = prepared["confirmation_token"]
    with caplog.at_level(logging.INFO, logger="mcp_o365.audit"):
        await run_email_move_confirm(
            "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
            approval=approval, confirmation_token=token, clock=clock,
        )
    await run_email_move_confirm(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, confirmation_token=token, clock=clock,
    )
    assert gc.count("move_message") == 1  # idempotente
    assert any(e["action"] == "email.move" for e in _audit_events(caplog))


# ============================ US-1.8 — ELIMINAR ============================


async def test_delete_soft_prepare_confirm(mapping, store, config, clock, caplog):
    _link(mapping, clock)
    gc = FakeGraphClient()
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)

    prepared = await run_email_delete_prepare(
        "subj-1", mapping=mapping, plane_b=pb, store=store, approval=approval,
        message_id="m1", clock=clock,
    )
    assert prepared["status"] == "pending_confirmation"
    assert gc.calls == []  # prepare não toca no Graph

    with caplog.at_level(logging.INFO, logger="mcp_o365.audit"):
        out = await run_email_delete_confirm(
            "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
            approval=approval, confirmation_token=prepared["confirmation_token"],
            clock=clock,
        )
    assert out["status"] == "done"
    # Soft delete = mover explicitamente para Itens Eliminados (não hard delete).
    assert gc.count("move_message") == 1
    assert gc.count("permanent_delete") == 0
    move = next(c for c in gc.calls if c[0] == "move_message")
    assert move[2]["destination_id"] == "deleteditems"
    assert any(e["action"] == "email.delete" for e in _audit_events(caplog))


async def test_delete_permanente_sem_confirmacao_reforcada_nao_apaga(
    mapping, store, config, clock
):
    """US-1.8 — eliminação permanente sem confirm_permanent -> error e NÃO apaga."""
    _link(mapping, clock)
    gc = FakeGraphClient()
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)

    prepared = await run_email_delete_prepare(
        "subj-1", mapping=mapping, plane_b=pb, store=store, approval=approval,
        message_id="m1", permanent=True, clock=clock,
    )
    assert prepared["requires_reinforced_confirmation"] is True
    token = prepared["confirmation_token"]

    # Sem confirm_permanent: recusa ANTES de consumir o token.
    blocked = await run_email_delete_confirm(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, confirmation_token=token, clock=clock,
    )
    assert blocked["status"] == "error"
    assert gc.count("permanent_delete") == 0 and gc.count("move_message") == 0

    # Com confirm_permanent=True: o mesmo token ainda é válido e apaga.
    ok = await run_email_delete_confirm(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, confirmation_token=token, confirm_permanent=True, clock=clock,
    )
    assert ok["status"] == "done"
    # Permanente = ação permanentDelete real (não soft/move).
    assert gc.count("permanent_delete") == 1
    assert gc.count("move_message") == 0


# ===================== REAUTH numa escrita =====================


async def test_escrita_reauth_quando_refresh_falha_nao_chama_graph(
    mapping, store, config, clock
):
    """Refresh falha (invalid_grant) no confirm -> reauth_required e Graph não é chamado."""
    # Token válido no prepare (não força refresh).
    _link(mapping, clock)
    gc = FakeGraphClient()
    approval = _approval(store, clock)
    prepared = await run_email_send_prepare(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock), store=store,
        approval=approval, to=["dest@example.com"], body="Olá", clock=clock,
    )
    # Entre o prepare e o confirm o token expira: o confirm tem de refrescar — e falha.
    store.update_account_tokens(
        subject="subj-1", account_id="acc-1", access_token="old-at",
        refresh_token="rt-1", expires_at=clock() - timedelta(minutes=5),
    )
    pb_bad = _plane_b(config, clock, refresh_result=INVALID_GRANT_RESPONSE)
    out = await run_email_send_confirm(
        "subj-1", mapping=mapping, plane_b=pb_bad, graph_client=gc, store=store,
        approval=approval, confirmation_token=prepared["confirmation_token"], clock=clock,
    )
    assert out["status"] == "reauth_required"
    assert gc.count("send_mail") == 0


# ===== Resiliência a 401/403 do Graph no confirm (cenário do encaminhamento reportado) =====


async def test_forward_confirm_recupera_de_401_transparente(mapping, store, config, clock):
    """O Graph recusa o token à 1ª (401) -> refresh forçado + retry -> envio concluído."""
    _link(mapping, clock)
    gc = FakeGraphClient(auth_fail={"forward": 1})  # falha uma vez, depois sucesso
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)

    prepared = await run_email_reply_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, message_id="m1", comment="FYI", mode="forward",
        to_recipients=["accounting@mobiweb.pt"], clock=clock,
    )
    confirmed = await run_email_reply_confirm(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, confirmation_token=prepared["confirmation_token"], clock=clock,
    )
    assert confirmed["status"] == "done"
    assert gc.count("forward") == 2  # 1ª falhou (401), 2ª com token renovado
    # Token renovado persistido no store.
    assert store.get_account("subj-1", "acc-1")["access_token"] == "graph-access-1"


async def test_forward_confirm_401_persistente_reauth_e_token_reutilizavel(
    mapping, store, config, clock
):
    """401 persistente -> reauth_required gracioso (sem erro cru) e token NÃO consumido."""
    _link(mapping, clock)
    gc = FakeGraphClient(auth_fail={"forward": 5})  # falha sempre
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)

    prepared = await run_email_reply_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, message_id="m1", comment="FYI", mode="forward",
        to_recipients=["accounting@mobiweb.pt"], clock=clock,
    )
    token = prepared["confirmation_token"]
    confirmed = await run_email_reply_confirm(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, confirmation_token=token, clock=clock,
    )
    assert confirmed["status"] == "reauth_required"
    # O token de confirmação continua por consumir -> pode repetir-se após o re-login.
    pending = store.get_pending_operation("subj-1", token)
    assert pending is not None and pending["consumed_at"] is None


# ===== US-1.4 — desambiguação reply vs reply_all (vários destinatários) =====


def _multi_recipient_msg() -> dict:
    return {
        "id": "m1",
        "toRecipients": ["a@x.com", "b@x.com"],
        "ccRecipients": ["c@x.com"],
    }


async def test_reply_com_varios_destinatarios_pede_clarificacao(mapping, store, config, clock):
    """mode='reply' + vários destinatários -> needs_clarification, sem criar token."""
    _link(mapping, clock)
    gc = FakeGraphClient(message=_multi_recipient_msg())
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)
    out = await run_email_reply_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, message_id="m1", comment="ok", mode="reply", clock=clock,
    )
    assert out["status"] == "needs_clarification"
    assert out["recipients_in_thread"] == 3
    assert "confirmation_token" not in out


async def test_reply_scope_confirmed_avanca_sem_perguntar(mapping, store, config, clock):
    """Com scope_confirmed=True, responde só ao remetente sem perguntar."""
    _link(mapping, clock)
    gc = FakeGraphClient(message=_multi_recipient_msg())
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)
    out = await run_email_reply_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, message_id="m1", comment="ok", mode="reply",
        scope_confirmed=True, clock=clock,
    )
    assert out["status"] == "pending_confirmation"


async def test_reply_unico_destinatario_nao_pergunta(mapping, store, config, clock):
    """Com um só destinatário, reply avança sem clarificação."""
    _link(mapping, clock)
    gc = FakeGraphClient(message={"id": "m1", "toRecipients": ["so-um@x.com"]})
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)
    out = await run_email_reply_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, message_id="m1", comment="ok", mode="reply", clock=clock,
    )
    assert out["status"] == "pending_confirmation"


# ==================== US-1.9 — AGENDAR ENVIO (T1-T11) ====================
#
# O relógio dos testes está fixo em 2026-06-01T12:00:00Z (FIXED_NOW). Por isso
# 2026-06-10T09:00:00Z é um instante futuro válido (>2 min, <1 ano).
_FUTURO_VALIDO = "2026-06-10T09:00:00Z"


def _deferred_prop(gc: FakeGraphClient) -> dict | None:
    """Devolve a única extended property do `singleValueExtendedProperties` da `message`
    passada ao `create_draft` (ou None se não houve create_draft)."""
    create = next((c for c in gc.calls if c[0] == "create_draft"), None)
    if create is None:
        return None
    message = create[1][1]  # args = (access_token, message)
    props = message.get("singleValueExtendedProperties") or []
    return props[0] if props else None


async def test_schedule_prepare_nao_toca_no_graph_T1(mapping, store, config, clock):
    """T1/T1b (AC1) — prepare devolve token e NÃO escreve no Graph.

    create_draft/send_draft a 0; get_mailbox_timezone <=1 (leitura best-effort do fuso)."""
    _link(mapping, clock)
    gc = FakeGraphClient(mailbox_timezone="Europe/Lisbon")
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)

    prepared = await run_email_schedule_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, to=["dest@mobiweb.pt"], body="Olá", subject_line="Teste",
        send_at=_FUTURO_VALIDO, clock=clock,
    )
    assert prepared["status"] == "pending_confirmation"
    assert prepared["confirmation_token"]
    assert prepared["recipients_count"] == 1
    assert prepared["large_attachments"] is False
    assert prepared["send_at_utc"] == "2026-06-10T09:00:00Z"
    # Nenhuma escrita no prepare.
    assert gc.count("create_draft") == 0
    assert gc.count("send_draft") == 0
    assert gc.count("send_mail") == 0
    assert gc.count("move_message") == 0
    # O fuso (leitura best-effort) pode ser consultado no máximo uma vez.
    assert gc.count("get_mailbox_timezone") <= 1


async def test_schedule_confirm_draft_send_com_propriedade_T2(
    mapping, store, config, clock, caplog
):
    """T2 (AC6) — confirm faz SEMPRE create_draft + send_draft (nunca send_mail), com a
    extended property de envio diferido na message do rascunho."""
    _link(mapping, clock)
    gc = FakeGraphClient(mailbox_timezone="Europe/Lisbon")
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)

    prepared = await run_email_schedule_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, to=["dest@mobiweb.pt"], body="Olá", subject_line="Teste",
        send_at=_FUTURO_VALIDO, clock=clock,
    )
    confirmed = await run_email_schedule_confirm(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, confirmation_token=prepared["confirmation_token"], clock=clock,
    )
    assert confirmed["status"] == "done"
    assert confirmed["message_id"] == "draft-1"  # id do rascunho diferido (para cancelar)
    # Diferido OBRIGA draft->send; nunca sendMail.
    assert gc.count("create_draft") == 1
    assert gc.count("send_draft") == 1
    assert gc.count("send_mail") == 0
    # A message do rascunho transporta a extended property PidTagDeferredSendTime.
    prop = _deferred_prop(gc)
    assert prop == {"id": _DEFERRED_SEND_PROP_ID, "value": "2026-06-10T09:00:00Z"}
    assert prop["id"] == "SystemTime 0x3FEF"


async def test_schedule_confirm_idempotente_nao_duplica_T3(mapping, store, config, clock):
    """T3 (AC1/idempotência) — replay do token -> idempotent_replay, sem segundo draft/send."""
    _link(mapping, clock)
    gc = FakeGraphClient()
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)

    prepared = await run_email_schedule_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, to=["dest@mobiweb.pt"], body="Olá",
        send_at=_FUTURO_VALIDO, clock=clock,
    )
    token = prepared["confirmation_token"]
    first = await run_email_schedule_confirm(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, confirmation_token=token, clock=clock,
    )
    second = await run_email_schedule_confirm(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, confirmation_token=token, clock=clock,
    )
    assert first["status"] == "done"
    assert second["status"] == "done"
    assert second["idempotent_replay"] is True
    # Nem duplo agendamento nem duplo envio.
    assert gc.count("create_draft") == 1
    assert gc.count("send_draft") == 1


async def test_schedule_normaliza_offset_para_utc_T4(mapping, store, config, clock):
    """T4 (AC3) — send_at com offset +01:00 -> propriedade em UTC (...08:00:00Z)."""
    _link(mapping, clock)
    gc = FakeGraphClient()
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)

    prepared = await run_email_schedule_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, to=["dest@mobiweb.pt"], body="Olá",
        send_at="2026-06-10T09:00:00+01:00", clock=clock,
    )
    # O instante normalizado já aparece no prepare.
    assert prepared["send_at_utc"] == "2026-06-10T08:00:00Z"

    await run_email_schedule_confirm(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, confirmation_token=prepared["confirmation_token"], clock=clock,
    )
    prop = _deferred_prop(gc)
    assert prop["value"] == "2026-06-10T08:00:00Z"  # 09:00 +01:00 -> 08:00 UTC


@pytest.mark.parametrize(
    "send_at",
    [
        "2026-05-01T09:00:00Z",          # passado
        "2026-06-01T12:00:30Z",          # < 2 min no futuro (now=12:00:00)
        "2030-01-01T00:00:00Z",          # > 1 ano
        "amanhã",                        # não-parseável
        "2026-06-10T09:00:00",           # sem offset (P3)
    ],
)
async def test_schedule_validacao_temporal_erro_sem_token_T5(
    mapping, store, config, clock, send_at
):
    """T5 (AC4) — passado / <2min / >1ano / não-parseável / sem-offset -> error SEM token,
    e sem escrever no Graph."""
    _link(mapping, clock)
    gc = FakeGraphClient()
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)

    out = await run_email_schedule_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, to=["dest@mobiweb.pt"], body="Olá",
        send_at=send_at, clock=clock,
    )
    assert out["status"] == "error"
    assert "confirmation_token" not in out
    assert gc.count("create_draft") == 0
    assert gc.count("send_draft") == 0


async def test_schedule_prepare_sem_to_erro_T6(mapping, store, config, clock):
    """T6 (AC2) — sem destinatários -> error."""
    _link(mapping, clock)
    out = await run_email_schedule_prepare(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock),
        graph_client=FakeGraphClient(), store=store, approval=_approval(store, clock),
        to=[], body="x", send_at=_FUTURO_VALIDO, clock=clock,
    )
    assert out["status"] == "error"
    assert "confirmation_token" not in out


async def test_schedule_anexo_grande_ordem_e_propriedade_T7(mapping, store, config, clock):
    """T7 (AC9) — anexo grande: create_draft -> upload session -> upload bytes -> send_draft,
    NESTA ordem, com a extended property no draft; send_mail==0."""
    _link(mapping, clock)
    gc = FakeGraphClient(draft={"id": "draft-1"})
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)

    big = {"name": "grande.bin", "contentBytes": "QUJD", "size": 4 * 1024 * 1024}
    prepared = await run_email_schedule_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, to=["dest@mobiweb.pt"], body="Olá", attachments=[big],
        send_at=_FUTURO_VALIDO, clock=clock,
    )
    assert prepared["large_attachments"] is True

    confirmed = await run_email_schedule_confirm(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, confirmation_token=prepared["confirmation_token"], clock=clock,
    )
    assert confirmed["status"] == "done"
    assert gc.count("create_draft") == 1
    assert gc.count("create_attachment_upload_session") == 1
    assert gc.count("upload_attachment_bytes") == 1
    assert gc.count("send_draft") == 1
    assert gc.count("send_mail") == 0
    # Ordem exata das chamadas de escrita.
    ordem = [c[0] for c in gc.calls if c[0] in (
        "create_draft", "create_attachment_upload_session",
        "upload_attachment_bytes", "send_draft",
    )]
    assert ordem == [
        "create_draft", "create_attachment_upload_session",
        "upload_attachment_bytes", "send_draft",
    ]
    # A extended property segue no rascunho mesmo com anexo grande.
    prop = _deferred_prop(gc)
    assert prop == {"id": _DEFERRED_SEND_PROP_ID, "value": "2026-06-10T09:00:00Z"}


async def test_schedule_fuso_indisponivel_degrada_para_utc_T8(mapping, store, config, clock):
    """T8 (AC3, degradação) — sem fuso do mailbox, o resumo declara UTC e a propriedade
    continua correta (o instante absoluto vem do send_at)."""
    _link(mapping, clock)
    gc = FakeGraphClient(mailbox_timezone=None)  # fuso indisponível
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)

    prepared = await run_email_schedule_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, to=["dest@mobiweb.pt"], body="Olá",
        send_at=_FUTURO_VALIDO, clock=clock,
    )
    assert prepared["status"] == "pending_confirmation"
    assert "UTC" in prepared["summary"]
    assert prepared["send_at_utc"] == "2026-06-10T09:00:00Z"


async def test_schedule_confirm_auditoria_so_metadados_T9(
    mapping, store, config, clock, caplog
):
    """T9 (AC8) — auditoria email.schedule: recipients_count, extra com
    large_attachments/send_at_utc/deferred, target=message_id; sem PII; um só subject_hash."""
    _link(mapping, clock)
    gc = FakeGraphClient(draft={"id": "draft-1"})
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)

    prepared = await run_email_schedule_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, to=["dest@mobiweb.pt", "outro@empresa.com"], body="Corpo secreto",
        subject_line="Assunto privado", send_at=_FUTURO_VALIDO, clock=clock,
    )
    with caplog.at_level(logging.INFO, logger="mcp_o365.audit"):
        await run_email_schedule_confirm(
            "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
            approval=approval, confirmation_token=prepared["confirmation_token"], clock=clock,
        )
    events = [e for e in _audit_events(caplog) if e.get("action") == "email.schedule"]
    assert len(events) == 1
    ev = events[0]
    assert ev["outcome"] == "success"
    assert ev["recipients_count"] == 2
    assert ev["target"] == "draft-1"
    assert ev["large_attachments"] is False
    assert ev["send_at_utc"] == "2026-06-10T09:00:00Z"
    assert ev["deferred"] is True
    # Regra A1: um único subject_hash (o de topo), sem PII de conteúdo.
    blob = repr(ev)
    assert "Corpo secreto" not in blob
    assert "Assunto privado" not in blob
    assert "@" not in blob
    assert sum(1 for k in ev if k == "subject_hash") == 1


async def test_schedule_prepare_reauth_sem_conta_T10(mapping, store, config, clock):
    """T10 (AC7) — prepare sem conta ligada -> reauth_required, sem escrever."""
    gc = FakeGraphClient()
    out = await run_email_schedule_prepare(
        "subj-sem-conta", mapping=mapping, plane_b=_plane_b(config, clock),
        graph_client=gc, store=store, approval=_approval(store, clock),
        to=["dest@mobiweb.pt"], body="x", send_at=_FUTURO_VALIDO, clock=clock,
    )
    assert out["status"] == "reauth_required"
    assert gc.count("create_draft") == 0


async def test_schedule_confirm_reauth_quando_refresh_falha_token_nao_consumido_T10b(
    mapping, store, config, clock
):
    """T10 (AC7) — refresh falha no confirm -> reauth_required, Graph não chamado e token
    NÃO consumido (repetível após re-login)."""
    _link(mapping, clock)
    gc = FakeGraphClient()
    approval = _approval(store, clock)
    prepared = await run_email_schedule_prepare(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock), graph_client=gc,
        store=store, approval=approval, to=["dest@mobiweb.pt"], body="Olá",
        send_at=_FUTURO_VALIDO, clock=clock,
    )
    token = prepared["confirmation_token"]
    # Token expira entre prepare e confirm; o refresh falha (invalid_grant).
    store.update_account_tokens(
        subject="subj-1", account_id="acc-1", access_token="old-at",
        refresh_token="rt-1", expires_at=clock() - timedelta(minutes=5),
    )
    pb_bad = _plane_b(config, clock, refresh_result=INVALID_GRANT_RESPONSE)
    out = await run_email_schedule_confirm(
        "subj-1", mapping=mapping, plane_b=pb_bad, graph_client=gc, store=store,
        approval=approval, confirmation_token=token, clock=clock,
    )
    assert out["status"] == "reauth_required"
    assert gc.count("create_draft") == 0
    assert gc.count("send_draft") == 0
    pending = store.get_pending_operation("subj-1", token)
    assert pending is not None and pending["consumed_at"] is None


async def test_schedule_aprendizagem_action_schedule_T11(
    mapping, store, config, clock, caplog
):
    """T11 (P5) — com opt-in ligado, o confirm regista record_action_event(action="schedule")
    -> audit learning.event_recorded com behavior_action="schedule"."""
    _link(mapping, clock)
    store.set_learning_opt_in("subj-1", True)
    gc = FakeGraphClient()
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)

    prepared = await run_email_schedule_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, to=["dest@mobiweb.pt"], body="Olá",
        send_at=_FUTURO_VALIDO,
        message_meta={"from": {"emailAddress": {"address": "chefe@mobiweb.pt"}}},
        clock=clock,
    )
    with caplog.at_level(logging.INFO, logger="mcp_o365.audit"):
        await run_email_schedule_confirm(
            "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
            approval=approval, confirmation_token=prepared["confirmation_token"], clock=clock,
        )
    learn = [
        e for e in _audit_events(caplog)
        if e.get("action") == "learning.event_recorded"
        and e.get("behavior_action") == "schedule"
    ]
    assert len(learn) == 1, (
        "esperava um evento de aprendizagem 'schedule' (contrato P5/T11)"
    )


async def test_schedule_aprendizagem_opt_out_nada_regista_T11b(
    mapping, store, config, clock, caplog
):
    """T11 (P5) — com opt-in DESLIGADO (default), nada é registado na aprendizagem."""
    _link(mapping, clock)
    gc = FakeGraphClient()
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)

    prepared = await run_email_schedule_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, to=["dest@mobiweb.pt"], body="Olá",
        send_at=_FUTURO_VALIDO, clock=clock,
    )
    with caplog.at_level(logging.INFO, logger="mcp_o365.audit"):
        await run_email_schedule_confirm(
            "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
            approval=approval, confirmation_token=prepared["confirmation_token"], clock=clock,
        )
    assert not [
        e for e in _audit_events(caplog)
        if e.get("action") == "learning.event_recorded"
    ]


# ==================== US-1.11 — CANCELAR AGENDAMENTO (T18-T24) ====================


def _deferred_msg(message_id: str = "draft-1", send_at: str = _FUTURO_VALIDO) -> dict:
    """Mensagem com a extended property de envio diferido (ainda agendada)."""
    return {
        "id": message_id,
        "subject": "Agendado",
        "singleValueExtendedProperties": [
            {"id": _DEFERRED_SEND_PROP_ID, "value": send_at}
        ],
    }


async def test_cancel_prepare_nao_escreve_T18(mapping, store, config, clock):
    """T18 (AC1) — cancel_prepare devolve token e NÃO escreve (move_message==0).
    get_message <=1 (leitura best-effort)."""
    _link(mapping, clock)
    gc = FakeGraphClient(message=_deferred_msg())
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)

    prepared = await run_email_schedule_cancel_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, message_id="draft-1", clock=clock,
    )
    assert prepared["status"] == "pending_confirmation"
    assert prepared["confirmation_token"]
    assert gc.count("move_message") == 0
    assert gc.count("get_message") <= 1


async def test_cancel_confirm_soft_delete_T19(mapping, store, config, clock, caplog):
    """T19 (AC4) — confirm faz soft delete (move para deleteditems); permanent_delete==0."""
    _link(mapping, clock)
    gc = FakeGraphClient(message=_deferred_msg(), moved={"id": "draft-1-eliminado"})
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)

    prepared = await run_email_schedule_cancel_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, message_id="draft-1", clock=clock,
    )
    with caplog.at_level(logging.INFO, logger="mcp_o365.audit"):
        out = await run_email_schedule_cancel_confirm(
            "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
            approval=approval, confirmation_token=prepared["confirmation_token"], clock=clock,
        )
    assert out["status"] == "done"
    assert gc.count("move_message") == 1
    assert gc.count("permanent_delete") == 0
    move = next(c for c in gc.calls if c[0] == "move_message")
    assert move[2]["destination_id"] == "deleteditems"


async def test_cancel_confirm_idempotente_T20(mapping, store, config, clock):
    """T20 (AC1/idempotência) — replay -> idempotent_replay; move_message continua a 1."""
    _link(mapping, clock)
    gc = FakeGraphClient(message=_deferred_msg())
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)

    prepared = await run_email_schedule_cancel_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, message_id="draft-1", clock=clock,
    )
    token = prepared["confirmation_token"]
    first = await run_email_schedule_cancel_confirm(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, confirmation_token=token, clock=clock,
    )
    second = await run_email_schedule_cancel_confirm(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, confirmation_token=token, clock=clock,
    )
    assert first["status"] == "done"
    assert second["idempotent_replay"] is True
    assert gc.count("move_message") == 1


async def test_cancel_prepare_ja_nao_diferido_erro_sem_token_T21(
    mapping, store, config, clock
):
    """T21 (AC3) — get_message sem a extended property -> error SEM token; move_message==0."""
    _link(mapping, clock)
    # Mensagem com a coleção presente mas sem a prop diferida (já enviado/cancelado).
    gc = FakeGraphClient(message={
        "id": "draft-1", "subject": "Já enviado",
        "singleValueExtendedProperties": [],
    })
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)

    out = await run_email_schedule_cancel_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, message_id="draft-1", clock=clock,
    )
    assert out["status"] == "error"
    assert "confirmation_token" not in out
    assert gc.count("move_message") == 0


async def test_cancel_prepare_get_message_degrada_emite_token_T22(
    mapping, store, config, clock
):
    """T22 (AC3, degradação) — leitura best-effort sem dados (get_message devolve {}) ->
    o prepare ainda emite token (resumo genérico)."""
    _link(mapping, clock)
    gc = FakeGraphClient(message={})  # sem subject nem prop -> degradação
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)

    prepared = await run_email_schedule_cancel_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, message_id="draft-1", clock=clock,
    )
    assert prepared["status"] == "pending_confirmation"
    assert prepared["confirmation_token"]
    assert gc.count("move_message") == 0


async def test_cancel_confirm_auditoria_T23(mapping, store, config, clock, caplog):
    """T23 (AC6) — auditoria email.schedule_cancel: target=message_id, extra permanent=False,
    sem PII, um só subject_hash."""
    _link(mapping, clock)
    gc = FakeGraphClient(message=_deferred_msg())
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)

    prepared = await run_email_schedule_cancel_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, message_id="draft-1", clock=clock,
    )
    with caplog.at_level(logging.INFO, logger="mcp_o365.audit"):
        await run_email_schedule_cancel_confirm(
            "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
            approval=approval, confirmation_token=prepared["confirmation_token"], clock=clock,
        )
    events = [e for e in _audit_events(caplog) if e.get("action") == "email.schedule_cancel"]
    assert len(events) == 1
    ev = events[0]
    assert ev["target"] == "draft-1"
    assert ev["permanent"] is False
    assert "@" not in repr(ev)
    assert sum(1 for k in ev if k == "subject_hash") == 1


async def test_cancel_confirm_reauth_token_nao_consumido_T24(mapping, store, config, clock):
    """T24 (AC5) — 401 persistente em move_message no confirm -> reauth_required gracioso e
    token NÃO consumido."""
    _link(mapping, clock)
    gc = FakeGraphClient(message=_deferred_msg(), auth_fail={"move_message": 5})
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)

    prepared = await run_email_schedule_cancel_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, message_id="draft-1", clock=clock,
    )
    token = prepared["confirmation_token"]
    out = await run_email_schedule_cancel_confirm(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, confirmation_token=token, clock=clock,
    )
    assert out["status"] == "reauth_required"
    pending = store.get_pending_operation("subj-1", token)
    assert pending is not None and pending["consumed_at"] is None


# ==================== TRANSVERSAIS (herdadas, análise §10) ====================


async def test_schedule_confirm_token_expirado_T_ttl(mapping, store, config, clock):
    """Transversal TTL — token expirado no confirm -> expired (sem escrever)."""
    _link(mapping, clock)
    gc = FakeGraphClient()
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)

    prepared = await run_email_schedule_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, to=["dest@mobiweb.pt"], body="Olá",
        send_at=_FUTURO_VALIDO, clock=clock,
    )
    # O TTL do ApprovalEngine é 300s; avança 301s -> token expira.
    clock.advance(301)
    out = await run_email_schedule_confirm(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, confirmation_token=prepared["confirmation_token"], clock=clock,
    )
    assert out["status"] == "expired"
    assert gc.count("create_draft") == 0
    assert gc.count("send_draft") == 0


async def test_schedule_confirm_token_de_outro_subject_T_isolamento(
    mapping, store, config, clock
):
    """Transversal isolamento — token de outro subject -> error, sem escrever."""
    _link(mapping, clock)
    mapping.link_account(
        subject="subj-2", access_token="valid-at-2", refresh_token="rt-2",
        expires_at=clock() + timedelta(hours=1), home_account_id="acc-2",
    )
    gc = FakeGraphClient()
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)

    prepared = await run_email_schedule_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, to=["dest@mobiweb.pt"], body="Olá",
        send_at=_FUTURO_VALIDO, clock=clock,
    )
    # subj-2 tenta usar o token preparado por subj-1.
    out = await run_email_schedule_confirm(
        "subj-2", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, confirmation_token=prepared["confirmation_token"], clock=clock,
    )
    assert out["status"] == "error"
    assert gc.count("create_draft") == 0
