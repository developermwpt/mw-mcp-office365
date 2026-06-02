"""Unit — GraphClient operações de email (T10/Fase 1) com HTTP mockado (respx).

Verifica a construção de query (`$search`+ConsistencyLevel, `$filter`), o mapeamento do
corpo, os endpoints de escrita, e a política de respostas (202/204 -> None) e retry 429
(Retry-After respeitado via sleeper fake, sem dormir).
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


@respx.mock
async def test_list_messages_monta_search_e_consistency_level():
    route = respx.get(f"{BASE}/me/messages").mock(
        return_value=httpx.Response(
            200,
            json={"value": [{"id": "m1", "subject": "Olá"}],
                  "@odata.nextLink": "https://next"},
        )
    )
    gc = _client()
    out = await gc.list_messages("tok", search="fatura", top=10)

    req = route.calls.last.request
    assert req.headers["ConsistencyLevel"] == "eventual"
    # `$` é URL-encoded como `%24`; o valor vem entre aspas (`"fatura"`).
    assert "%24search=%22fatura%22" in str(req.url)
    assert out["messages"][0]["id"] == "m1"
    assert out["next"] == "https://next"


@respx.mock
async def test_list_messages_monta_filter():
    route = respx.get(f"{BASE}/me/messages").mock(
        return_value=httpx.Response(200, json={"value": []})
    )
    gc = _client()
    out = await gc.list_messages("tok", filter_query="from/emailAddress/address eq 'a@b.com'")
    req = route.calls.last.request
    assert "%24filter=" in str(req.url) or "$filter=" in str(req.url)
    # Sem search -> sem ConsistencyLevel.
    assert "ConsistencyLevel" not in req.headers
    assert out["next"] is None


@respx.mock
async def test_list_messages_de_uma_pasta():
    route = respx.get(f"{BASE}/me/mailFolders/inbox/messages").mock(
        return_value=httpx.Response(200, json={"value": []})
    )
    gc = _client()
    await gc.list_messages("tok", folder="inbox")
    assert route.called


@respx.mock
async def test_get_message_mapeia_body():
    respx.get(f"{BASE}/me/messages/m1").mock(
        return_value=httpx.Response(
            200,
            json={
                "id": "m1", "subject": "Assunto",
                "from": {"emailAddress": {"address": "a@b.com"}},
                "toRecipients": [{"emailAddress": {"address": "c@d.com"}}],
                "body": {"contentType": "html", "content": "<p>oi</p>"},
                "hasAttachments": True,
            },
        )
    )
    gc = _client()
    msg = await gc.get_message("tok", "m1")
    assert msg["from"] == "a@b.com"
    assert msg["toRecipients"] == ["c@d.com"]
    assert msg["body"]["contentType"] == "html"
    assert msg["body"]["content"] == "<p>oi</p>"
    assert msg["hasAttachments"] is True


@respx.mock
async def test_send_mail_post_com_save_to_sent_items():
    route = respx.post(f"{BASE}/me/sendMail").mock(return_value=httpx.Response(202))
    gc = _client()
    out = await gc.send_mail("tok", message={"subject": "x"}, save_to_sent_items=True)
    assert out is None  # 202 -> None
    body = route.calls.last.request.read()
    assert b'"saveToSentItems": true' in body or b'"saveToSentItems":true' in body
    assert b'"subject"' in body


@respx.mock
async def test_reply_e_reply_all_batem_nos_endpoints():
    r1 = respx.post(f"{BASE}/me/messages/m1/reply").mock(return_value=httpx.Response(202))
    r2 = respx.post(f"{BASE}/me/messages/m1/replyAll").mock(return_value=httpx.Response(202))
    gc = _client()
    await gc.reply("tok", "m1", comment="ok")
    await gc.reply("tok", "m1", comment="ok", reply_all=True)
    assert r1.called
    assert r2.called


@respx.mock
async def test_forward_endpoint_e_recipients():
    route = respx.post(f"{BASE}/me/messages/m1/forward").mock(
        return_value=httpx.Response(202)
    )
    gc = _client()
    await gc.forward("tok", "m1", comment="fyi", to_recipients=["x@y.com"])
    body = route.calls.last.request.read()
    assert b"x@y.com" in body
    assert b'"comment"' in body


@respx.mock
async def test_move_message_post_move():
    route = respx.post(f"{BASE}/me/messages/m1/move").mock(
        return_value=httpx.Response(201, json={"id": "m1-novo"})
    )
    gc = _client()
    out = await gc.move_message("tok", "m1", destination_id="archive")
    assert out["id"] == "m1-novo"
    body = route.calls.last.request.read()
    assert b"archive" in body


@respx.mock
async def test_delete_message_delete_204():
    route = respx.delete(f"{BASE}/me/messages/m1").mock(return_value=httpx.Response(204))
    gc = _client()
    out = await gc.delete_message("tok", "m1")
    assert out is None
    assert route.called


@respx.mock
async def test_list_e_get_attachment():
    respx.get(f"{BASE}/me/messages/m1/attachments").mock(
        return_value=httpx.Response(
            200, json={"value": [{"id": "a1", "name": "f.pdf", "size": 10}]}
        )
    )
    respx.get(f"{BASE}/me/messages/m1/attachments/a1").mock(
        return_value=httpx.Response(
            200, json={"id": "a1", "name": "f.pdf", "contentBytes": "QUJD"}
        )
    )
    gc = _client()
    lst = await gc.list_attachments("tok", "m1")
    assert lst[0]["name"] == "f.pdf"
    att = await gc.get_attachment("tok", "m1", "a1")
    assert att["contentBytes"] == "QUJD"


@respx.mock
async def test_429_respeita_retry_after_e_repete():
    slept: list[float] = []

    async def fake_sleep(s: float) -> None:
        slept.append(s)

    route = respx.post(f"{BASE}/me/sendMail")
    route.side_effect = [
        httpx.Response(429, headers={"Retry-After": "3"}),
        httpx.Response(202),
    ]
    gc = GraphClient(httpx.AsyncClient(), sleeper=fake_sleep)
    out = await gc.send_mail("tok", message={"subject": "x"})
    assert out is None
    assert slept == [3.0]  # respeitou Retry-After (via sleeper, sem dormir) e repetiu
    assert len(route.calls) == 2
