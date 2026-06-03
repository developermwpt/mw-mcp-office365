"""Unit — GraphClient operações de calendário (Fase 2) com HTTP mockado (respx).

Cobre: mapeadores `_map_event_summary`/`_map_event_detail`/`_is_recurring`; injeção do
header `Prefer: outlook.timezone` nas leituras; construção do body do `getSchedule`; e o
roteamento de `respond_event` para o path certo (accept/decline/tentativelyAccept).
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


def test_map_event_summary_campos():
    e = {
        "id": "evt-1",
        "subject": "Reunião",
        "start": {"dateTime": "2026-06-10T10:00:00", "timeZone": "GMT Standard Time"},
        "end": {"dateTime": "2026-06-10T11:00:00", "timeZone": "GMT Standard Time"},
        "location": {"displayName": "Sala 1"},
        "organizer": {"emailAddress": {"address": "chefe@example.com"}},
        "isOnlineMeeting": True,
        "onlineMeeting": {"joinUrl": "https://teams/join"},
        "isAllDay": False,
        "bodyPreview": "olá",
        "type": "singleInstance",
    }
    out = GraphClient._map_event_summary(e)
    assert out["id"] == "evt-1"
    assert out["start"]["timeZone"] == "GMT Standard Time"
    assert out["location"] == "Sala 1"
    assert out["organizer"] == "chefe@example.com"
    assert out["isOnlineMeeting"] is True
    assert out["joinUrl"] == "https://teams/join"
    assert out["isRecurring"] is False
    assert out["bodyPreview"] == "olá"


def test_map_event_detail_attendees_e_body():
    e = {
        "id": "evt-2",
        "subject": "Sync",
        "attendees": [
            {
                "emailAddress": {"address": "a@b.com", "name": "Ana"},
                "type": "required",
                "status": {"response": "accepted"},
            }
        ],
        "responseStatus": {"response": "tentativelyAccepted"},
        "body": {"contentType": "html", "content": "<p>oi</p>"},
        "webLink": "https://web/evt-2",
        "type": "occurrence",
        "seriesMasterId": "series-1",
    }
    out = GraphClient._map_event_detail(e)
    assert out["attendees"][0] == {
        "email": "a@b.com", "name": "Ana", "type": "required",
        "responseStatus": "accepted",
    }
    assert out["responseStatus"] == "tentativelyAccepted"
    assert out["body"]["content"] == "<p>oi</p>"  # CRU; sanitizado na tool
    assert out["webLink"] == "https://web/evt-2"
    assert out["type"] == "occurrence"
    assert out["isRecurring"] is True  # tem seriesMasterId


def test_is_recurring():
    assert GraphClient._is_recurring({"type": "singleInstance"}) is False
    assert GraphClient._is_recurring({"type": "occurrence"}) is True
    assert GraphClient._is_recurring({"type": "seriesMaster"}) is True
    assert GraphClient._is_recurring({"type": "exception"}) is True
    assert GraphClient._is_recurring({"seriesMasterId": "s1"}) is True
    assert GraphClient._is_recurring({}) is False


# ===================== leitura + header Prefer =====================


@respx.mock
async def test_mailbox_timezone():
    respx.get(f"{BASE}/me/mailboxSettings").mock(
        return_value=httpx.Response(200, json={"timeZone": "GMT Standard Time"})
    )
    gc = _client()
    assert await gc.get_mailbox_timezone("tok") == "GMT Standard Time"


@respx.mock
async def test_list_calendar_view_injeta_prefer_e_params():
    route = respx.get(f"{BASE}/me/calendarView").mock(
        return_value=httpx.Response(
            200,
            json={"value": [{"id": "evt-1", "subject": "X"}],
                  "@odata.nextLink": "https://next-cal"},
        )
    )
    gc = _client()
    out = await gc.list_calendar_view(
        "tok", start="2026-06-10T00:00:00Z", end="2026-06-11T00:00:00Z",
        top=50, prefer_timezone="GMT Standard Time",
    )
    req = route.calls.last.request
    assert req.headers["Prefer"] == 'outlook.timezone="GMT Standard Time"'
    assert "startDateTime=" in str(req.url)
    assert "%24orderby=start%2FdateTime" in str(req.url) or "$orderby" in str(req.url)
    assert out["events"][0]["id"] == "evt-1"
    assert out["next"] == "https://next-cal"


@respx.mock
async def test_list_calendar_view_sem_tz_nao_envia_prefer():
    route = respx.get(f"{BASE}/me/calendarView").mock(
        return_value=httpx.Response(200, json={"value": []})
    )
    gc = _client()
    await gc.list_calendar_view("tok", start="a", end="b")
    assert "prefer" not in {k.lower() for k in route.calls.last.request.headers}


@respx.mock
async def test_list_calendar_view_next_segue_link_absoluto_com_prefer():
    route = respx.get("https://next-cal-2").mock(
        return_value=httpx.Response(200, json={"value": [{"id": "evt-9"}], })
    )
    gc = _client()
    out = await gc.list_calendar_view_next(
        "tok", "https://next-cal-2", prefer_timezone="W. Europe Standard Time"
    )
    assert route.calls.last.request.headers["Prefer"] == (
        'outlook.timezone="W. Europe Standard Time"'
    )
    assert out["events"][0]["id"] == "evt-9"


@respx.mock
async def test_get_event_mapeia_detalhe():
    respx.get(f"{BASE}/me/events/evt-1").mock(
        return_value=httpx.Response(
            200,
            json={
                "id": "evt-1", "subject": "Reunião",
                "organizer": {"emailAddress": {"address": "org@x.com"}},
                "responseStatus": {"response": "accepted"},
                "body": {"contentType": "text", "content": "corpo"},
            },
        )
    )
    gc = _client()
    out = await gc.get_event("tok", "evt-1")
    assert out["organizer"] == "org@x.com"
    assert out["responseStatus"] == "accepted"


# ===================== getSchedule =====================


@respx.mock
async def test_get_schedule_monta_body_e_mapeia():
    route = respx.post(f"{BASE}/me/calendar/getSchedule").mock(
        return_value=httpx.Response(
            200,
            json={
                "value": [
                    {
                        "scheduleId": "a@x.com",
                        "availabilityView": "0022",
                        "scheduleItems": [
                            {
                                "status": "busy",
                                "start": {"dateTime": "2026-06-10T10:00:00"},
                                "end": {"dateTime": "2026-06-10T11:00:00"},
                            }
                        ],
                    }
                ]
            },
        )
    )
    gc = _client()
    out = await gc.get_schedule(
        "tok", schedules=["a@x.com"], start="2026-06-10T00:00:00",
        end="2026-06-11T00:00:00", interval_minutes=30,
        prefer_timezone="GMT Standard Time",
    )
    body = route.calls.last.request.read()
    assert b'"schedules"' in body and b"a@x.com" in body
    assert b'"availabilityViewInterval": 30' in body or b'"availabilityViewInterval":30' in body
    assert out[0]["email"] == "a@x.com"
    assert out[0]["availabilityView"] == "0022"
    assert out[0]["scheduleItems"][0]["status"] == "busy"
    assert out[0]["scheduleItems"][0]["start"] == "2026-06-10T10:00:00"


# ===================== escrita =====================


@respx.mock
async def test_create_event_post_e_mapeia():
    route = respx.post(f"{BASE}/me/events").mock(
        return_value=httpx.Response(
            201, json={"id": "evt-novo", "webLink": "https://web/evt-novo"}
        )
    )
    gc = _client()
    out = await gc.create_event("tok", event={"subject": "X"})
    assert b'"subject"' in route.calls.last.request.read()
    assert out["id"] == "evt-novo"
    assert out["webLink"] == "https://web/evt-novo"


@respx.mock
async def test_update_event_patch():
    route = respx.patch(f"{BASE}/me/events/evt-1").mock(
        return_value=httpx.Response(200, json={"id": "evt-1"})
    )
    gc = _client()
    await gc.update_event("tok", "evt-1", changes={"subject": "Novo"})
    assert b"Novo" in route.calls.last.request.read()


@respx.mock
async def test_cancel_event_post_204():
    route = respx.post(f"{BASE}/me/events/evt-1/cancel").mock(
        return_value=httpx.Response(202)
    )
    gc = _client()
    out = await gc.cancel_event("tok", "evt-1", comment="adiado")
    assert out is None
    assert b"adiado" in route.calls.last.request.read()


@respx.mock
async def test_respond_event_mapeia_para_path_certo():
    accept = respx.post(f"{BASE}/me/events/evt-1/accept").mock(
        return_value=httpx.Response(202)
    )
    decline = respx.post(f"{BASE}/me/events/evt-1/decline").mock(
        return_value=httpx.Response(202)
    )
    tentative = respx.post(f"{BASE}/me/events/evt-1/tentativelyAccept").mock(
        return_value=httpx.Response(202)
    )
    gc = _client()
    await gc.respond_event("tok", "evt-1", response="accept", comment="ok")
    await gc.respond_event("tok", "evt-1", response="decline")
    await gc.respond_event("tok", "evt-1", response="tentativelyAccept")
    assert accept.called and decline.called and tentative.called
    body = accept.calls.last.request.read()
    assert b'"sendResponse": true' in body or b'"sendResponse":true' in body
