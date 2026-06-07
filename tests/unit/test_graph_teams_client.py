"""Unit — GraphClient operações de Teams (Fase 3) com HTTP mockado (respx).

Cobre: mapeadores `_map_chat_summary`/`_map_chat_message`/`_chat_from`/`_map_chat_member`;
construção dos pedidos (`$expand`/`$top`/`$orderby`) de `list_chats`/`get_chat`/
`list_chat_messages`; body de `send_chat_message` (text e html) e de `create_one_on_one_chat`
(`chatType=oneOnOne` + `user@odata.bind`); paginação por `@odata.nextLink` absoluto.
"""

from __future__ import annotations

import httpx
import respx

from mcp_o365.graph.client import GraphClient

BASE = "https://graph.microsoft.com/v1.0"


async def _noop_sleep(_s: float) -> None:
    return None


def _client() -> GraphClient:
    return GraphClient(httpx.AsyncClient(), sleeper=_noop_sleep)


# ===================== mapeadores =====================


def test_map_chat_summary_oneonone_sem_topico():
    c = {
        "id": "chat-1",
        "chatType": "oneOnOne",
        "topic": None,
        "members": [
            {"displayName": "Ana", "email": "ana@x.com", "userId": "u-ana"},
            {"displayName": "Eu", "email": "eu@x.com", "userId": "u-eu"},
        ],
        "lastUpdatedDateTime": "2026-06-06T10:00:00Z",
        "lastMessagePreview": {"body": {"content": "olá"}},
    }
    out = GraphClient._map_chat_summary(c)
    assert out["id"] == "chat-1"
    assert out["chat_type"] == "oneOnOne"
    assert out["topic"] is None
    assert out["members"][0] == {"name": "Ana", "email": "ana@x.com", "aad_user_id": "u-ana"}
    assert out["last_updated"] == "2026-06-06T10:00:00Z"
    assert out["last_message_preview"] == "olá"  # CRU; sanitizado na tool


def test_map_chat_summary_grupo_com_topico_e_membro_sem_email():
    c = {
        "id": "chat-g",
        "chatType": "group",
        "topic": "Projeto X",
        "members": [
            {"displayName": "Bruno", "userId": "u-bruno"},  # sem email -> fallback aad_user_id
        ],
    }
    out = GraphClient._map_chat_summary(c)
    assert out["chat_type"] == "group"
    assert out["topic"] == "Projeto X"
    assert out["members"][0] == {"name": "Bruno", "email": None, "aad_user_id": "u-bruno"}
    # lastMessagePreview ausente -> None (R6), sem rebentar.
    assert out["last_message_preview"] is None


def test_map_chat_message_normal_e_sistema():
    normal = {
        "id": "m-1",
        "from": {"user": {"displayName": "Ana", "email": "ana@x.com"}},
        "createdDateTime": "2026-06-06T10:00:00Z",
        "messageType": "message",
        "body": {"contentType": "html", "content": "<p>oi</p>"},
        "attachments": [{"id": "a1"}],
    }
    out = GraphClient._map_chat_message(normal)
    assert out["is_system"] is False
    assert out["from"] == {"name": "Ana", "email": "ana@x.com"}
    assert out["body"]["content"] == "<p>oi</p>"  # CRU; sanitizado na tool
    assert out["attachments_count"] == 1

    system = {"id": "m-2", "from": None, "messageType": "systemEventMessage",
              "body": {"contentType": "text", "content": "Ana entrou"}}
    out2 = GraphClient._map_chat_message(system)
    assert out2["is_system"] is True  # D8
    assert out2["from"] is None
    assert out2["attachments_count"] == 0


def test_chat_from_aplicacao_e_nulo():
    assert GraphClient._chat_from({"from": None}) is None
    assert GraphClient._chat_from({"from": {"application": {"displayName": "Bot"}}}) is None
    assert GraphClient._chat_from({"from": {"user": {"displayName": "Zé"}}}) == {
        "name": "Zé", "email": None,
    }


# ===================== construção dos pedidos =====================


@respx.mock
async def test_list_chats_monta_expand_top_sem_orderby_e_ordena_client_side():
    # Regressão: /me/chats NÃO suporta $orderby (Graph 400 BadRequest). Confirmamos
    # que NÃO o enviamos e que a ordenação por last_updated desc é feita client-side.
    route = respx.get(f"{BASE}/me/chats").mock(
        return_value=httpx.Response(
            200,
            json={"value": [
                {"id": "chat-antigo", "chatType": "oneOnOne",
                 "lastUpdatedDateTime": "2026-06-01T10:00:00Z"},
                {"id": "chat-recente", "chatType": "group",
                 "lastUpdatedDateTime": "2026-06-06T10:00:00Z"},
            ], "@odata.nextLink": "https://next-chats"},
        )
    )
    out = await _client().list_chats("tok", top=50)
    url = str(route.calls.last.request.url)
    assert "members" in url and "lastMessagePreview" in url
    assert "%24top=50" in url or "$top=50" in url
    assert "orderby" not in url.lower() and "lastUpdatedDateTime" not in url
    # ordenado desc client-side: o mais recente primeiro
    assert [c["id"] for c in out["chats"]] == ["chat-recente", "chat-antigo"]
    assert out["next"] == "https://next-chats"


@respx.mock
async def test_get_chat_monta_expand_members():
    route = respx.get(f"{BASE}/me/chats/chat-1").mock(
        return_value=httpx.Response(
            200,
            json={"id": "chat-1", "chatType": "group", "topic": "T",
                  "members": [{"displayName": "Ana", "email": "ana@x.com"}]},
        )
    )
    out = await _client().get_chat("tok", "chat-1")
    url = str(route.calls.last.request.url)
    assert "members" in url
    assert out["chat_type"] == "group"
    assert out["members"][0]["email"] == "ana@x.com"


@respx.mock
async def test_list_chat_messages_monta_top_sem_orderby_e_ordena_client_side():
    # Regressão: /chats/{id}/messages NÃO suporta $orderby; garantimos a ordem
    # (createdDateTime desc) client-side em vez de a pedir ao Graph.
    route = respx.get(f"{BASE}/me/chats/chat-1/messages").mock(
        return_value=httpx.Response(
            200,
            json={"value": [
                {"id": "m-antiga", "messageType": "message",
                 "createdDateTime": "2026-06-01T09:00:00Z"},
                {"id": "m-recente", "messageType": "message",
                 "createdDateTime": "2026-06-06T09:00:00Z"},
            ], "@odata.nextLink": "https://next-msgs"},
        )
    )
    out = await _client().list_chat_messages("tok", "chat-1", top=25)
    url = str(route.calls.last.request.url)
    assert "%24top=25" in url or "$top=25" in url
    assert "orderby" not in url.lower() and "createdDateTime" not in url
    assert [m["id"] for m in out["messages"]] == ["m-recente", "m-antiga"]
    assert out["next"] == "https://next-msgs"


@respx.mock
async def test_send_chat_message_text_e_html():
    route = respx.post(f"{BASE}/me/chats/chat-1/messages").mock(
        return_value=httpx.Response(201, json={"id": "m-new", "messageType": "message"})
    )
    gc = _client()
    out = await gc.send_chat_message("tok", "chat-1", content="olá", content_type="text")
    body = route.calls.last.request.read()
    assert b'"contentType": "text"' in body or b'"contentType":"text"' in body
    assert b"ol\\u00e1" in body or "olá".encode() in body
    assert out["id"] == "m-new"

    await gc.send_chat_message("tok", "chat-1", content="<b>x</b>", content_type="html")
    body2 = route.calls.last.request.read()
    assert b'"contentType": "html"' in body2 or b'"contentType":"html"' in body2


@respx.mock
async def test_create_one_on_one_chat_monta_body():
    route = respx.post(f"{BASE}/chats").mock(
        return_value=httpx.Response(201, json={"id": "chat-novo", "chatType": "oneOnOne"})
    )
    out = await _client().create_one_on_one_chat(
        "tok", member_emails=["eu@x.com", "ana@x.com"]
    )
    body = route.calls.last.request.read()
    assert b'"chatType": "oneOnOne"' in body or b'"chatType":"oneOnOne"' in body
    assert b"user@odata.bind" in body
    assert b"ana@x.com" in body and b"eu@x.com" in body
    assert out["id"] == "chat-novo"


# ===================== paginação (nextLink absoluto) =====================


@respx.mock
async def test_list_chats_next_segue_link_absoluto():
    respx.get("https://next-chats-2").mock(
        return_value=httpx.Response(200, json={"value": [{"id": "chat-9"}]})
    )
    out = await _client().list_chats_next("tok", "https://next-chats-2")
    assert out["chats"][0]["id"] == "chat-9"
    assert out["next"] is None


@respx.mock
async def test_list_chat_messages_next_segue_link_absoluto():
    respx.get("https://next-msgs-2").mock(
        return_value=httpx.Response(
            200, json={"value": [{"id": "m-9", "messageType": "message"}],
                       "@odata.nextLink": "https://next-msgs-3"},
        )
    )
    out = await _client().list_chat_messages_next("tok", "https://next-msgs-2")
    assert out["messages"][0]["id"] == "m-9"
    assert out["next"] == "https://next-msgs-3"
