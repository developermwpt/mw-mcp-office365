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

from mcp_o365.approval.engine import ApprovalEngine
from mcp_o365.auth.plane_b import PlaneB
from mcp_o365.tools.email import (
    run_email_delete_confirm,
    run_email_delete_prepare,
    run_email_move_confirm,
    run_email_move_prepare,
    run_email_reply_confirm,
    run_email_reply_prepare,
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
        "subj-1", mapping=mapping, plane_b=pb, store=store, approval=approval,
        message_id="m1", comment="Obrigado", mode="reply", clock=clock,
    )
    assert prepared["status"] == "pending_confirmation"
    assert gc.calls == []  # prepare não toca no Graph

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
        "subj-1", mapping=mapping, plane_b=pb, store=store, approval=approval,
        message_id="m1", comment="Para todos", mode="reply_all", clock=clock,
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
        "subj-1", mapping=mapping, plane_b=pb, store=store, approval=approval,
        message_id="m1", comment="FYI", mode="forward",
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
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock), store=store,
        approval=_approval(store, clock), message_id="m1", comment="x",
        mode="forward", clock=clock,
    )
    assert out["status"] == "error"


async def test_reply_mode_invalido_erro(mapping, store, config, clock):
    _link(mapping, clock)
    out = await run_email_reply_prepare(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock), store=store,
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
    assert gc.count("delete_message") == 1
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
    assert gc.count("delete_message") == 0

    # Com confirm_permanent=True: o mesmo token ainda é válido e apaga.
    ok = await run_email_delete_confirm(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, confirmation_token=token, confirm_permanent=True, clock=clock,
    )
    assert ok["status"] == "done"
    assert gc.count("delete_message") == 1


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
        "subj-1", mapping=mapping, plane_b=pb, store=store, approval=approval,
        message_id="m1", comment="FYI", mode="forward",
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
        "subj-1", mapping=mapping, plane_b=pb, store=store, approval=approval,
        message_id="m1", comment="FYI", mode="forward",
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
