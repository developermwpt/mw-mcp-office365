"""Unit — GraphClient operações de email (T10/Fase 1) com HTTP mockado (respx).

Verifica a construção de query (`$search`+ConsistencyLevel, `$filter`), o mapeamento do
corpo, os endpoints de escrita, e a política de respostas (202/204 -> None) e retry 429
(Retry-After respeitado via sleeper fake, sem dormir).
"""

from __future__ import annotations

import httpx
import respx

from mcp_o365.graph.client import GraphClient, GraphError

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


@respx.mock
async def test_upload_attachment_bytes_em_chunks_sem_bearer():
    """US-1.6 — bytes carregados em chunks com Content-Range; uploadUrl sem Authorization."""
    url = "https://upload.example/session"
    route = respx.put(url).mock(return_value=httpx.Response(202))
    gc = _client()
    await gc.upload_attachment_bytes(url, b"0123456789", chunk_size=4)

    ranges = [c.request.headers["Content-Range"] for c in route.calls]
    assert ranges == ["bytes 0-3/10", "bytes 4-7/10", "bytes 8-9/10"]
    # A uploadUrl é pré-autenticada — não se envia o Bearer.
    assert "authorization" not in {k.lower() for k in route.calls[0].request.headers}
    # Os bytes reconstruídos batem certo com o original.
    sent = b"".join(c.request.content for c in route.calls)
    assert sent == b"0123456789"


@respx.mock
async def test_upload_attachment_bytes_erro_levanta_grapherror():
    url = "https://upload.example/session"
    respx.put(url).mock(return_value=httpx.Response(500, text="boom"))
    gc = _client()
    try:
        await gc.upload_attachment_bytes(url, b"abc", chunk_size=4)
        raise AssertionError("devia ter levantado GraphError")
    except GraphError:
        pass


# ==================== US-1.9/1.10/1.11 — AGENDAMENTO ====================


@respx.mock
async def test_get_message_com_expand_expoe_extended_property():
    """US-1.11/P7 — get_message(expand=...) propaga $expand e expõe
    singleValueExtendedProperties no resultado."""
    route = respx.get(f"{BASE}/me/messages/d1").mock(
        return_value=httpx.Response(
            200,
            json={
                "id": "d1", "subject": "Agendado",
                "singleValueExtendedProperties": [
                    {"id": "SystemTime 0x3FEF", "value": "2026-06-10T09:00:00Z"}
                ],
            },
        )
    )
    gc = _client()
    msg = await gc.get_message(
        "tok", "d1",
        expand="singleValueExtendedProperties($filter=id eq 'SystemTime 0x3FEF')",
    )
    url = str(route.calls.last.request.url)
    assert "%24expand=" in url or "$expand=" in url
    assert msg["singleValueExtendedProperties"][0]["value"] == "2026-06-10T09:00:00Z"


@respx.mock
async def test_get_message_sem_expand_retrocompativel():
    """Sem expand, o get_message comporta-se como antes (sem $expand, sem a coleção)."""
    route = respx.get(f"{BASE}/me/messages/m1").mock(
        return_value=httpx.Response(
            200, json={"id": "m1", "subject": "Normal"}
        )
    )
    gc = _client()
    msg = await gc.get_message("tok", "m1")
    url = str(route.calls.last.request.url)
    assert "expand" not in url
    assert "singleValueExtendedProperties" not in msg


@respx.mock
async def test_list_deferred_drafts_query_filter_expand_select():
    """US-1.10 — list_deferred_drafts monta $filter (presença), $expand (valor), $select e
    $top sobre /me/mailFolders/drafts/messages; mapeia os rascunhos."""
    route = respx.get(f"{BASE}/me/mailFolders/drafts/messages").mock(
        return_value=httpx.Response(
            200,
            json={
                "value": [
                    {
                        "id": "d1", "subject": "Agendado",
                        "toRecipients": [
                            {"emailAddress": {"address": "a@mobiweb.pt"}},
                            {"emailAddress": {"address": "b@empresa.com"}},
                        ],
                        "singleValueExtendedProperties": [
                            {"id": "SystemTime 0x3FEF", "value": "2026-06-10T09:00:00Z"}
                        ],
                    }
                ],
                "@odata.nextLink": "https://next-drafts",
            },
        )
    )
    gc = _client()
    out = await gc.list_deferred_drafts("tok", prop_id="SystemTime 0x3FEF", top=25)

    url = str(route.calls.last.request.url)
    # $filter testa a PRESENÇA da propriedade; $expand traz o valor; $select minimiza.
    assert "any(ep" in url or "any%28ep" in url
    assert "SystemTime+0x3FEF" in url or "SystemTime%200x3FEF" in url
    assert "%24expand=" in url or "$expand=" in url
    assert "%24top=25" in url or "$top=25" in url

    assert out["next"] == "https://next-drafts"
    draft = out["drafts"][0]
    assert draft["id"] == "d1"
    assert draft["subject"] == "Agendado"
    assert draft["to"] == ["a@mobiweb.pt", "b@empresa.com"]
    assert draft["deferred_send_at"] == "2026-06-10T09:00:00Z"


@respx.mock
async def test_list_deferred_drafts_sem_prop_mapeia_none():
    """O mapeamento devolve deferred_send_at=None quando a prop não está presente."""
    respx.get(f"{BASE}/me/mailFolders/drafts/messages").mock(
        return_value=httpx.Response(
            200,
            json={"value": [{"id": "d2", "subject": "Sem prop", "toRecipients": []}]},
        )
    )
    gc = _client()
    out = await gc.list_deferred_drafts("tok", prop_id="SystemTime 0x3FEF")
    assert out["drafts"][0]["deferred_send_at"] is None
    assert out["drafts"][0]["to"] == []
    assert out["next"] is None
