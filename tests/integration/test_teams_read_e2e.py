"""Integração — tools de LEITURA de Teams end-to-end (US-3.1, US-3.2).

Store e mapping reais; Graph mockado por `FakeGraphClient`; Entra por FakeMsalApp. As
leituras não exigem aprovação. Prova-se: filtro client-side (D2) por tópico E por membro,
`members` só nome+email, preview sanitizado + `content_is_untrusted`, clamp do `top` (D4),
`has_more`/`next_link` (D5), `page_token` -> `list_chat_messages_next`, `is_system` (D8),
corpo HTML sanitizado e reauth graciosa.
"""

from __future__ import annotations

from datetime import timedelta

from mcp_o365.auth.plane_b import PlaneB
from mcp_o365.tools import teams
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


def _chats_payload() -> dict:
    # O FakeGraphClient substitui o GraphClient: devolve já o formato MAPEADO
    # (_map_chat_summary), tal como o fake do calendário devolve eventos mapeados.
    return {
        "chats": [
            {
                "id": "chat-11",
                "chat_type": "oneOnOne",
                "topic": None,
                "members": [
                    {"name": "Ana Silva", "email": "ana@mobiweb.pt", "aad_user_id": "u-a"},
                    {"name": "Eu", "email": "eu@mobiweb.pt", "aad_user_id": "u-eu"},
                ],
                "last_updated": "2026-06-06T09:00:00Z",
                "last_message_preview": "olá<script>roubar()</script>",
            },
            {
                "id": "chat-22",
                "chat_type": "group",
                "topic": "Projeto Moomenti",
                "members": [
                    {"name": "Bruno", "email": "bruno@cliente.com", "aad_user_id": "u-b"},
                ],
                "last_updated": "2026-06-06T10:00:00Z",
                "last_message_preview": "estado?",
            },
        ],
        "next": None,
    }


# ============================ US-3.1 — LISTAR CHATS ============================


async def test_list_chats_simples(mapping, store, config, clock):
    """US-3.1 — sem filtro: devolve todos; members só nome+email+aad; content_is_untrusted."""
    _link(mapping, clock)
    gc = FakeGraphClient(chats=_chats_payload())
    out = await teams.run_teams_list_chats(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock),
        graph_client=gc, store=store, clock=clock,
    )
    assert out["status"] == "ok"
    assert out["count"] == 2
    assert out["content_is_untrusted"] is True
    assert out["has_more"] is False
    # Ordenado por last_updated desc (defensivo): o grupo (10:00) vem primeiro.
    assert out["chats"][0]["id"] == "chat-22"
    member = out["chats"][1]["members"][0]
    assert set(member.keys()) == {"name", "email", "aad_user_id"}
    assert gc.count("list_chats") == 1
    assert gc.count("list_chats_next") == 0


async def test_list_chats_preview_sanitizado(mapping, store, config, clock):
    """US-3.1 — last_message_preview HTML é sanitizado (não-confiável)."""
    _link(mapping, clock)
    gc = FakeGraphClient(chats=_chats_payload())
    out = await teams.run_teams_list_chats(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock),
        graph_client=gc, store=store, clock=clock,
    )
    oneonone = next(c for c in out["chats"] if c["id"] == "chat-11")
    assert "<script>" not in (oneonone["last_message_preview"] or "")
    assert "roubar" not in oneonone["last_message_preview"]
    assert "olá" in oneonone["last_message_preview"]


async def test_list_chats_filtro_por_topico(mapping, store, config, clock):
    """US-3.1 — filter_text por tópico -> só o grupo."""
    _link(mapping, clock)
    gc = FakeGraphClient(chats=_chats_payload())
    out = await teams.run_teams_list_chats(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock),
        graph_client=gc, store=store, filter_text="moomenti", clock=clock,
    )
    assert out["count"] == 1
    assert out["chats"][0]["id"] == "chat-22"


async def test_list_chats_filtro_por_membro(mapping, store, config, clock):
    """US-3.1 — filter_text por nome/email de membro -> só os chats com esse membro."""
    _link(mapping, clock)
    gc = FakeGraphClient(chats=_chats_payload())
    out_nome = await teams.run_teams_list_chats(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock),
        graph_client=gc, store=store, filter_text="ana silva", clock=clock,
    )
    assert [c["id"] for c in out_nome["chats"]] == ["chat-11"]
    gc2 = FakeGraphClient(chats=_chats_payload())
    out_email = await teams.run_teams_list_chats(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock),
        graph_client=gc2, store=store, filter_text="bruno@cliente.com", clock=clock,
    )
    assert [c["id"] for c in out_email["chats"]] == ["chat-22"]


async def test_list_chats_filtro_pagina_ate_satisfazer(mapping, store, config, clock):
    """US-3.1 — com filtro e next: pagina via list_chats_next (D2 client-side)."""
    _link(mapping, clock)
    page1 = {
        "chats": [
            {"id": "c1", "chat_type": "group", "topic": "Outro",
             "members": [], "last_updated": "2026-06-01T00:00:00Z",
             "last_message_preview": None},
        ],
        "next": "https://chats-p2",
    }
    page2 = {
        "chats": [
            {"id": "c2", "chat_type": "group", "topic": "Alvo",
             "members": [], "last_updated": "2026-06-02T00:00:00Z",
             "last_message_preview": None},
        ],
        "next": None,
    }
    gc = FakeGraphClient(chats=page1, next_chat_pages=[page2])
    out = await teams.run_teams_list_chats(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock),
        graph_client=gc, store=store, filter_text="alvo", clock=clock,
    )
    assert [c["id"] for c in out["chats"]] == ["c2"]
    assert gc.count("list_chats_next") == 1


async def test_list_chats_sem_conta_reauth(mapping, store, config, clock):
    """US-3.1 — sem conta ligada -> reauth_required, Graph não tocado."""
    gc = FakeGraphClient(chats=_chats_payload())
    out = await teams.run_teams_list_chats(
        "subj-sem-conta", mapping=mapping, plane_b=_plane_b(config, clock),
        graph_client=gc, store=store, clock=clock,
    )
    assert out["status"] == "reauth_required"
    assert gc.calls == []


# ============================ US-3.2 — LER MENSAGENS ============================


def _messages_payload(next_link=None) -> dict:
    # Formato MAPEADO (_map_chat_message): o fake substitui o GraphClient.
    return {
        "messages": [
            {
                "id": "m-1",
                "from": {"name": "Ana", "email": "ana@x.com"},
                "created": "2026-06-06T10:00:00Z",
                "message_type": "message",
                "is_system": False,
                "body": {"contentType": "html", "content": "<b>olá</b><script>x</script>"},
                "attachments_count": 0,
            },
            {
                "id": "m-2",
                "from": None,
                "created": "2026-06-06T09:00:00Z",
                "message_type": "systemEventMessage",
                "is_system": True,
                "body": {"contentType": "text", "content": "Ana entrou no chat"},
                "attachments_count": 0,
            },
        ],
        "next": next_link,
    }


async def test_read_messages_default_top_e_sanitiza(mapping, store, config, clock):
    """US-3.2 — top default 25; corpo HTML sanitizado; is_system (D8); content_is_untrusted."""
    _link(mapping, clock)
    gc = FakeGraphClient(chat_messages=_messages_payload())
    out = await teams.run_teams_read_messages(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock),
        graph_client=gc, store=store, chat_id="chat-11", clock=clock,
    )
    assert out["status"] == "ok"
    assert out["chat_id"] == "chat-11"
    assert out["count"] == 2
    assert out["content_is_untrusted"] is True
    assert out["has_more"] is False
    # top default = 25 (D4).
    call = next(c for c in gc.calls if c[0] == "list_chat_messages")
    assert call[2]["top"] == 25
    # corpo HTML sanitizado.
    m1 = next(m for m in out["messages"] if m["id"] == "m-1")
    assert "<script>" not in m1["body"]["content"]
    # mensagem de sistema marcada (D8).
    m2 = next(m for m in out["messages"] if m["id"] == "m-2")
    assert m2["is_system"] is True
    assert m2["from"] is None


async def test_read_messages_clamp_top_a_50(mapping, store, config, clock):
    """US-3.2 — top=999 -> clamp a 50 (D4)."""
    _link(mapping, clock)
    gc = FakeGraphClient(chat_messages=_messages_payload())
    await teams.run_teams_read_messages(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock),
        graph_client=gc, store=store, chat_id="chat-11", top=999, clock=clock,
    )
    call = next(c for c in gc.calls if c[0] == "list_chat_messages")
    assert call[2]["top"] == 50


async def test_read_messages_has_more_e_next_link(mapping, store, config, clock):
    """US-3.2 — has_more=true + next_link quando há next."""
    _link(mapping, clock)
    gc = FakeGraphClient(chat_messages=_messages_payload(next_link="https://msgs-p2"))
    out = await teams.run_teams_read_messages(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock),
        graph_client=gc, store=store, chat_id="chat-11", clock=clock,
    )
    assert out["has_more"] is True
    assert out["next_link"] == "https://msgs-p2"


async def test_read_messages_page_token_usa_next(mapping, store, config, clock):
    """US-3.2 — page_token chama list_chat_messages_next e NÃO list_chat_messages (D5)."""
    _link(mapping, clock)
    gc = FakeGraphClient(
        next_message_pages=[{"messages": [{"id": "m-old", "message_type": "message",
                                           "is_system": False,
                                           "body": {"contentType": "text", "content": "x"}}],
                             "next": None}],
    )
    out = await teams.run_teams_read_messages(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock),
        graph_client=gc, store=store, chat_id="chat-11",
        page_token="https://msgs-old", clock=clock,
    )
    assert out["status"] == "ok"
    assert out["messages"][0]["id"] == "m-old"
    assert gc.count("list_chat_messages_next") == 1
    assert gc.count("list_chat_messages") == 0


async def test_read_messages_sem_chat_id_erro(mapping, store, config, clock):
    """US-3.2 — chat_id em falta -> error, Graph não tocado."""
    _link(mapping, clock)
    gc = FakeGraphClient(chat_messages=_messages_payload())
    out = await teams.run_teams_read_messages(
        "subj-1", mapping=mapping, plane_b=_plane_b(config, clock),
        graph_client=gc, store=store, chat_id="", clock=clock,
    )
    assert out["status"] == "error"
    assert gc.calls == []


async def test_read_messages_reauth(mapping, store, config, clock):
    """US-3.2 — sem conta -> reauth_required."""
    gc = FakeGraphClient(chat_messages=_messages_payload())
    out = await teams.run_teams_read_messages(
        "subj-sem-conta", mapping=mapping, plane_b=_plane_b(config, clock),
        graph_client=gc, store=store, chat_id="chat-11", clock=clock,
    )
    assert out["status"] == "reauth_required"
    assert gc.calls == []
