"""Integração — tools de ESCRITA de Teams end-to-end (US-3.3, US-3.4, US-3.5).

Cada escrita segue o par prepare/confirm. Invariantes verificadas com o `FakeGraphClient` a
contar chamadas:
- `prepare` NÃO chama o método de escrita (`send_chat_message`/`create_one_on_one_chat` a 0);
- o prepare de envio só LÊ via `get_chat` (A2) — count=1, sem escrita;
- `confirm` chama o método certo 1× e regista auditoria (event=audit);
- replay do token é idempotente (não duplica o envio/criação);
- auditoria `teams.send` com extra={chat_type, body_type} SEM subject_hash em extra (A1) —
  mas o subject_hash da IDENTIDADE de topo continua presente;
- chat 1:1 existente -> ok SEM token (create a 0); inexistente -> pending_confirmation;
- reauth graciosa no prepare e no confirm (token não consumido).
"""

from __future__ import annotations

import logging
from datetime import timedelta

from mcp_o365.approval.engine import ApprovalEngine
from mcp_o365.auth.plane_b import PlaneB
from mcp_o365.tools import teams
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


def _group_chat() -> dict:
    # Formato MAPEADO devolvido pelo get_chat (substitui o GraphClient real).
    return {
        "id": "chat-22",
        "chat_type": "group",
        "topic": "Projeto Moomenti",
        "members": [
            {"name": "Ana", "email": "ana@mobiweb.pt", "aad_user_id": "u-a"},
            {"name": "Bruno", "email": "bruno@cliente.com", "aad_user_id": "u-b"},
            {"name": "Eu", "email": "eu@mobiweb.pt", "aad_user_id": "u-eu"},
        ],
        "last_updated": "2026-06-06T10:00:00Z",
        "last_message_preview": None,
    }


# ============================ US-3.3 / US-3.5 — ENVIAR ============================


def _one_on_one_chat() -> dict:
    # 1:1 mapeado: o próprio (eu@mobiweb.pt) + a destinatária.
    return {
        "id": "chat-11",
        "chat_type": "oneOnOne",
        "topic": None,
        "members": [
            {"name": "Vera Martins", "email": "vera.martins@mobiweb.pt", "aad_user_id": "u-v"},
            {"name": "Eu", "email": "eu@mobiweb.pt", "aad_user_id": "u-eu"},
        ],
        "last_updated": "2026-06-06T10:00:00Z",
        "last_message_preview": None,
    }


async def test_send_prepare_le_get_chat_nao_escreve(mapping, store, config, clock):
    """US-3.3 — prepare lê via get_chat (count=1, A2) e NÃO envia (send a 0). Num GRUPO o
    resumo declara N participantes EXCLUINDO o próprio (recipients_count semântico) + domínios."""
    _link(mapping, clock)
    # me == "Eu" do grupo -> o emissor é excluído da contagem (3 membros -> 2 destinatários).
    gc = FakeGraphClient(chat=_group_chat(), me={"userPrincipalName": "eu@mobiweb.pt"})
    prepared = await teams.run_teams_send_message_prepare(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock), graph_client=gc,
        store=store, approval=_approval(store, clock), chat_id="chat-22",
        body="olá equipa", clock=clock,
    )
    assert prepared["status"] == "pending_confirmation"
    assert prepared["operation"] == "teams.send"
    # Emissor excluído: 3 membros - o próprio = 2 destinatários.
    assert prepared["recipients_count"] == 2
    assert "grupo" in prepared["summary"]
    assert "2 participante(s)" in prepared["summary"]
    assert "Projeto Moomenti" in prepared["summary"]
    assert "mobiweb.pt" in prepared["summary"] and "cliente.com" in prepared["summary"]
    assert "formato: text" in prepared["summary"]
    # A2: prepare lê via get_chat, NÃO envia.
    assert gc.count("get_chat") == 1
    assert gc.count("send_chat_message") == 0


async def test_send_prepare_1a1_nomeia_destinatario_e_conta_1(mapping, store, config, clock):
    """US-3.3/3.5 — num 1:1 o emissor é excluído (recipients_count=1) e o resumo NOMEIA a
    pessoa (barreira concreta), em vez de contar participantes."""
    _link(mapping, clock)
    gc = FakeGraphClient(
        chat=_one_on_one_chat(), me={"userPrincipalName": "eu@mobiweb.pt"}
    )
    prepared = await teams.run_teams_send_message_prepare(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock), graph_client=gc,
        store=store, approval=_approval(store, clock), chat_id="chat-11",
        body="olá Vera", clock=clock,
    )
    assert prepared["status"] == "pending_confirmation"
    assert prepared["recipients_count"] == 1
    assert "Vera Martins" in prepared["summary"]
    assert "vera.martins@mobiweb.pt" in prepared["summary"]
    assert "participante(s)" not in prepared["summary"]  # 1:1 nomeia, não conta
    assert gc.count("send_chat_message") == 0


async def test_send_prepare_get_chat_degrada_mas_emite_token(mapping, store, config, clock):
    """US-3.3 — get_chat a falhar (não-auth) -> resumo degrada sem detalhes mas ainda emite
    token; ainda assim não escreve."""
    _link(mapping, clock)
    gc = FakeGraphClient(chat=_group_chat(), auth_fail={"get_chat": 99})
    prepared = await teams.run_teams_send_message_prepare(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock), graph_client=gc,
        store=store, approval=_approval(store, clock), chat_id="chat-22",
        body="olá", clock=clock,
    )
    assert prepared["status"] == "pending_confirmation"
    assert prepared["recipients_count"] == 0
    assert "indisponíveis" in prepared["summary"]
    assert gc.count("send_chat_message") == 0


async def test_send_html_aceite(mapping, store, config, clock):
    """US-3.3 — body_type='html' aceite; o resumo declara o formato."""
    _link(mapping, clock)
    gc = FakeGraphClient(chat=_group_chat())
    prepared = await teams.run_teams_send_message_prepare(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock), graph_client=gc,
        store=store, approval=_approval(store, clock), chat_id="chat-22",
        body="<b>oi</b>", body_type="html", clock=clock,
    )
    assert prepared["status"] == "pending_confirmation"
    assert "formato: html" in prepared["summary"]


async def test_send_body_type_invalido_erro(mapping, store, config, clock):
    """US-3.3 — body_type='xml' -> error, sem token e sem leitura/escrita."""
    _link(mapping, clock)
    gc = FakeGraphClient(chat=_group_chat())
    out = await teams.run_teams_send_message_prepare(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock), graph_client=gc,
        store=store, approval=_approval(store, clock), chat_id="chat-22",
        body="oi", body_type="xml", clock=clock,
    )
    assert out["status"] == "error"
    assert "confirmation_token" not in out
    assert gc.calls == []


async def test_send_body_demasiado_longo_erro(mapping, store, config, clock):
    """US-3.3 — body acima do teto (D10) -> error sem token (validado antes de qualquer IO)."""
    _link(mapping, clock)
    gc = FakeGraphClient(chat=_group_chat())
    out = await teams.run_teams_send_message_prepare(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock), graph_client=gc,
        store=store, approval=_approval(store, clock), chat_id="chat-22",
        body="x" * 28001, clock=clock,
    )
    assert out["status"] == "error"
    assert "demasiado longa" in out["message"]
    assert gc.calls == []


async def test_send_falta_campos_erro(mapping, store, config, clock):
    """US-3.3 — chat_id/body em falta -> error, Graph não tocado."""
    _link(mapping, clock)
    gc = FakeGraphClient(chat=_group_chat())
    out = await teams.run_teams_send_message_prepare(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock), graph_client=gc,
        store=store, approval=_approval(store, clock), chat_id="", body="oi", clock=clock,
    )
    assert out["status"] == "error"
    assert gc.calls == []


async def test_send_confirm_envia_e_audita_sem_subject_hash_em_extra(
    mapping, store, config, clock, caplog
):
    """US-3.3 — confirm envia 1× e audita teams.send com extra={chat_type, body_type}, SEM
    subject_hash em extra (A1); o subject_hash de IDENTIDADE (topo) continua presente."""
    _link(mapping, clock)
    gc = FakeGraphClient(
        chat=_group_chat(), sent_message={"id": "msg-99"},
        me={"userPrincipalName": "eu@mobiweb.pt"},
    )
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)
    prepared = await teams.run_teams_send_message_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, chat_id="chat-22", body="olá", clock=clock,
    )
    with caplog.at_level(logging.INFO, logger="mcp_o365.audit"):
        confirmed = await teams.run_teams_send_message_confirm(
            "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
            approval=approval, confirmation_token=prepared["confirmation_token"], clock=clock,
        )
    assert confirmed["status"] == "done"
    assert confirmed["chat_id"] == "chat-22"
    assert confirmed["message_id"] == "msg-99"
    assert gc.count("send_chat_message") == 1
    # O payload de envio chegou com o corpo/format certos.
    call = next(c for c in gc.calls if c[0] == "send_chat_message")
    assert call[2]["content"] == "olá"
    assert call[2]["content_type"] == "text"
    # Auditoria só-metadados (A1).
    audit = next(e for e in _audit_events(caplog) if e["action"] == "teams.send")
    assert audit["outcome"] == "success"
    assert audit["target"] == "chat-22"
    assert audit["recipients_count"] == 2  # emissor excluído (3 membros - o próprio)
    assert audit["chat_type"] == "group"
    assert audit["body_type"] == "text"
    # subject_hash de identidade presente (topo) e não foi sobrescrito por extra.
    assert audit["subject_hash"]
    assert "olá" not in str(audit)  # nunca o texto da mensagem


async def test_send_confirm_idempotente(mapping, store, config, clock):
    """US-3.3 — replay do token não re-envia; send_chat_message fica a 1 (anti-duplicação)."""
    _link(mapping, clock)
    gc = FakeGraphClient(chat=_group_chat())
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)
    prepared = await teams.run_teams_send_message_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, chat_id="chat-22", body="olá", clock=clock,
    )
    token = prepared["confirmation_token"]
    first = await teams.run_teams_send_message_confirm(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, confirmation_token=token, clock=clock,
    )
    second = await teams.run_teams_send_message_confirm(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, confirmation_token=token, clock=clock,
    )
    assert first["status"] == "done"
    assert second["idempotent_replay"] is True
    assert gc.count("send_chat_message") == 1


async def test_send_confirm_reauth_nao_envia_nem_consome_token(
    mapping, store, config, clock
):
    """US-3.3 — refresh falha no confirm -> reauth_required, sem envio e token reutilizável."""
    _link(mapping, clock)
    gc = FakeGraphClient(chat=_group_chat())
    approval = _approval(store, clock)
    prepared = await teams.run_teams_send_message_prepare(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock), graph_client=gc,
        store=store, approval=approval, chat_id="chat-22", body="olá", clock=clock,
    )
    token = prepared["confirmation_token"]
    store.update_account_tokens(
        subject="subj-1", account_id="acc-1", access_token="old-at",
        refresh_token="rt-1", expires_at=clock() - timedelta(minutes=5),
    )
    pb_bad = _plane_b(config, clock, refresh_result=INVALID_GRANT_RESPONSE)
    out = await teams.run_teams_send_message_confirm(
        "subj-1", mapping=mapping, plane_b=pb_bad, graph_client=gc, store=store,
        approval=approval, confirmation_token=token, clock=clock,
    )
    assert out["status"] == "reauth_required"
    assert gc.count("send_chat_message") == 0
    pending = store.get_pending_operation("subj-1", token)
    assert pending is not None and pending["consumed_at"] is None


async def test_send_prepare_reauth(mapping, store, config, clock):
    """US-3.3 — sem conta no prepare -> reauth_required, Graph não tocado."""
    gc = FakeGraphClient(chat=_group_chat())
    out = await teams.run_teams_send_message_prepare(
        "subj-sem-conta", mapping=mapping, plane_b=_plane_b(config, clock), graph_client=gc,
        store=store, approval=_approval(store, clock), chat_id="chat-22", body="olá",
        clock=clock,
    )
    assert out["status"] == "reauth_required"
    assert gc.count("send_chat_message") == 0


# US-3.5 — responder = enviar no mesmo chat_id (reusa o par de US-3.3).
async def test_us35_responder_reusa_send_no_mesmo_chat(mapping, store, config, clock):
    """US-3.5 — 'responder' num chat 1:1 é enviar nova mensagem no mesmo chat_id (D7)."""
    _link(mapping, clock)
    one_on_one = {
        "id": "chat-11", "chat_type": "oneOnOne", "topic": None,
        "members": [
            {"name": "Ana", "email": "ana@mobiweb.pt", "aad_user_id": "u-a"},
            {"name": "Eu", "email": "eu@mobiweb.pt", "aad_user_id": "u-eu"},
        ],
        "last_updated": "2026-06-06T10:00:00Z", "last_message_preview": None,
    }
    gc = FakeGraphClient(chat=one_on_one, me={"userPrincipalName": "eu@mobiweb.pt"})
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)
    prepared = await teams.run_teams_send_message_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, chat_id="chat-11", body="resposta", clock=clock,
    )
    # 1:1 nomeia a pessoa (emissor excluído) e conta 1 destinatário.
    assert "Ana" in prepared["summary"]
    assert prepared["recipients_count"] == 1
    await teams.run_teams_send_message_confirm(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, confirmation_token=prepared["confirmation_token"], clock=clock,
    )
    call = next(c for c in gc.calls if c[0] == "send_chat_message")
    assert call[1][1] == "chat-11"  # mesmo chat_id


# ============================ US-3.4 — OBTER/CRIAR 1:1 ============================


def _chats_with_one_on_one(other_email="ana@mobiweb.pt") -> dict:
    return {
        "chats": [
            {
                "id": "chat-11",
                "chat_type": "oneOnOne",
                "topic": None,
                "members": [
                    {"name": "Ana", "email": other_email, "aad_user_id": "u-a"},
                    {"name": "Eu", "email": "subj@example.com", "aad_user_id": "u-eu"},
                ],
                "last_updated": "2026-06-06T10:00:00Z",
                "last_message_preview": None,
            },
        ],
        "next": None,
    }


async def test_get_or_create_chat_existente_ok_sem_token(mapping, store, config, clock):
    """US-3.4 — chat 1:1 já existe -> status='ok' + chat_id, sem token; create a 0."""
    _link(mapping, clock)
    gc = FakeGraphClient(
        me={"userPrincipalName": "subj@example.com"},
        chats=_chats_with_one_on_one(),
    )
    out = await teams.run_teams_get_or_create_one_on_one_chat_prepare(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock), graph_client=gc,
        store=store, approval=_approval(store, clock), member_email="ana@mobiweb.pt",
        clock=clock,
    )
    assert out["status"] == "ok"
    assert out["chat_id"] == "chat-11"
    assert out["is_new_chat"] is False
    assert "confirmation_token" not in out
    assert gc.count("create_one_on_one_chat") == 0


async def test_get_or_create_chat_inexistente_pending(mapping, store, config, clock):
    """US-3.4 — chat inexistente -> pending_confirmation (token); create ainda a 0."""
    _link(mapping, clock)
    gc = FakeGraphClient(
        me={"userPrincipalName": "subj@example.com"},
        chats={"chats": [], "next": None},
    )
    out = await teams.run_teams_get_or_create_one_on_one_chat_prepare(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock), graph_client=gc,
        store=store, approval=_approval(store, clock), member_email="novo@x.com", clock=clock,
    )
    assert out["status"] == "pending_confirmation"
    assert out["operation"] == "teams.chat_create"
    assert "INICIAR" in out["summary"] and "novo@x.com" in out["summary"]
    assert gc.count("create_one_on_one_chat") == 0


async def test_get_or_create_chat_membro_sem_email_cria(mapping, store, config, clock):
    """US-3.4 (A5) — membro sem email -> sem match -> segue para criação (pending)."""
    _link(mapping, clock)
    chats = {
        "chats": [
            {"id": "chat-x", "chat_type": "oneOnOne", "topic": None,
             "members": [
                 {"name": "Ana", "email": None, "aad_user_id": "u-a"},
                 {"name": "Eu", "email": "subj@example.com", "aad_user_id": "u-eu"},
             ],
             "last_updated": "2026-06-06T10:00:00Z", "last_message_preview": None},
        ],
        "next": None,
    }
    gc = FakeGraphClient(me={"userPrincipalName": "subj@example.com"}, chats=chats)
    out = await teams.run_teams_get_or_create_one_on_one_chat_prepare(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock), graph_client=gc,
        store=store, approval=_approval(store, clock), member_email="ana@mobiweb.pt",
        clock=clock,
    )
    assert out["status"] == "pending_confirmation"
    assert gc.count("create_one_on_one_chat") == 0


async def test_get_or_create_chat_confirm_cria_e_audita(mapping, store, config, clock, caplog):
    """US-3.4 — confirm cria 1× e audita teams.chat_create com {chat_type, is_new_chat}."""
    _link(mapping, clock)
    gc = FakeGraphClient(
        me={"userPrincipalName": "subj@example.com"},
        chats={"chats": [], "next": None},
        created_chat={"id": "chat-criado", "chat_type": "oneOnOne"},
    )
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)
    prepared = await teams.run_teams_get_or_create_one_on_one_chat_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, member_email="novo@x.com", clock=clock,
    )
    with caplog.at_level(logging.INFO, logger="mcp_o365.audit"):
        confirmed = await teams.run_teams_get_or_create_one_on_one_chat_confirm(
            "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
            approval=approval, confirmation_token=prepared["confirmation_token"], clock=clock,
        )
    assert confirmed["status"] == "done"
    assert confirmed["chat_id"] == "chat-criado"
    assert confirmed["is_new_chat"] is True
    assert gc.count("create_one_on_one_chat") == 1
    # O body de criação levou os dois emails.
    call = next(c for c in gc.calls if c[0] == "create_one_on_one_chat")
    assert call[2]["member_emails"] == ["subj@example.com", "novo@x.com"]
    audit = next(e for e in _audit_events(caplog) if e["action"] == "teams.chat_create")
    assert audit["chat_type"] == "oneOnOne"
    assert audit["is_new_chat"] is True
    assert audit["target"] == "chat-criado"


async def test_get_or_create_chat_confirm_idempotente(mapping, store, config, clock):
    """US-3.4 — replay não re-cria; create_one_on_one_chat fica a 1."""
    _link(mapping, clock)
    gc = FakeGraphClient(
        me={"userPrincipalName": "subj@example.com"},
        chats={"chats": [], "next": None},
    )
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)
    prepared = await teams.run_teams_get_or_create_one_on_one_chat_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, member_email="novo@x.com", clock=clock,
    )
    token = prepared["confirmation_token"]
    await teams.run_teams_get_or_create_one_on_one_chat_confirm(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, confirmation_token=token, clock=clock,
    )
    second = await teams.run_teams_get_or_create_one_on_one_chat_confirm(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, confirmation_token=token, clock=clock,
    )
    assert second["idempotent_replay"] is True
    assert gc.count("create_one_on_one_chat") == 1


async def test_get_or_create_chat_sem_member_email_erro(mapping, store, config, clock):
    """US-3.4 — member_email em falta -> error, Graph não tocado."""
    _link(mapping, clock)
    gc = FakeGraphClient(me={"userPrincipalName": "subj@example.com"})
    out = await teams.run_teams_get_or_create_one_on_one_chat_prepare(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock), graph_client=gc,
        store=store, approval=_approval(store, clock), member_email="", clock=clock,
    )
    assert out["status"] == "error"
    assert gc.calls == []


async def test_get_or_create_chat_prepare_reauth(mapping, store, config, clock):
    """US-3.4 — sem conta no prepare -> reauth_required, sem criação."""
    gc = FakeGraphClient(me={"userPrincipalName": "subj@example.com"})
    out = await teams.run_teams_get_or_create_one_on_one_chat_prepare(
        "subj-sem-conta", mapping=mapping, plane_b=_plane_b(config, clock), graph_client=gc,
        store=store, approval=_approval(store, clock), member_email="x@y.com", clock=clock,
    )
    assert out["status"] == "reauth_required"
    assert gc.count("create_one_on_one_chat") == 0


# ===================== TRANSVERSAIS =====================


async def test_send_token_expirado_devolve_expired(mapping, store, config, clock):
    """Transversal — TTL expirado no confirm -> expired; sem envio."""
    _link(mapping, clock)
    gc = FakeGraphClient(chat=_group_chat())
    approval = ApprovalEngine(store, clock=clock, ttl_seconds=60)
    pb = _plane_b(config, clock)
    prepared = await teams.run_teams_send_message_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, chat_id="chat-22", body="olá", clock=clock,
    )
    clock.advance(120)
    out = await teams.run_teams_send_message_confirm(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, confirmation_token=prepared["confirmation_token"], clock=clock,
    )
    assert out["status"] == "expired"
    assert gc.count("send_chat_message") == 0


async def test_send_token_de_outro_subject_rejeitado(mapping, store, config, clock):
    """Transversal — token de outro subject -> error; sem envio."""
    _link(mapping, clock)
    mapping.link_account(
        subject="subj-2", access_token="valid-at", refresh_token="rt-1",
        expires_at=clock() + timedelta(hours=1), home_account_id="acc-2",
    )
    gc = FakeGraphClient(chat=_group_chat())
    approval = _approval(store, clock)
    pb = _plane_b(config, clock)
    prepared = await teams.run_teams_send_message_prepare(
        "subj-1", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, chat_id="chat-22", body="olá", clock=clock,
    )
    out = await teams.run_teams_send_message_confirm(
        "subj-2", mapping=mapping, plane_b=pb, graph_client=gc, store=store,
        approval=approval, confirmation_token=prepared["confirmation_token"], clock=clock,
    )
    assert out["status"] == "error"
    assert gc.count("send_chat_message") == 0
