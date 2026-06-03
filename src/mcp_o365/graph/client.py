"""T10 — Wrapper fino sobre o Microsoft Graph.

Na Fase 0 só expunha `me()` (`GET /me`). A Fase 1 generaliza-o num `_request` único (com a
mesma lógica de retry 429 / 401-403 / >=400) e acrescenta as operações de email (ler,
pesquisar, anexos, enviar, responder, reencaminhar, mover, eliminar).

Não conhece tokens nem store — recebe sempre o access token pronto (resolvido a montante
por `resolve_access_token`). O cliente HTTP e o `sleeper` são injetados para testes
determinísticos (sem rede nem `sleep` reais).
"""

from __future__ import annotations

import asyncio
from collections.abc import Awaitable, Callable
from typing import Any

import httpx

from ..auth.errors import UpstreamAuthError

DEFAULT_GRAPH_BASE = "https://graph.microsoft.com/v1.0"
_MAX_RETRIES_429 = 3
# Tamanho de chunk para upload de anexos grandes: múltiplo de 320 KiB (requisito Graph).
_UPLOAD_CHUNK_SIZE = 320 * 1024 * 10  # ~3,2 MB


async def _real_sleeper(seconds: float) -> None:
    await asyncio.sleep(seconds)


class GraphError(Exception):
    """Erro genérico do Graph (não relacionado com auth)."""


def recipients(addresses: list[str]) -> list[dict]:
    """Converte uma lista de emails no formato de destinatários do Graph."""
    return [{"emailAddress": {"address": addr}} for addr in addresses if addr]


class GraphClient:
    """Cliente HTTP mínimo para o Microsoft Graph."""

    def __init__(
        self,
        http: httpx.AsyncClient,
        base_url: str = DEFAULT_GRAPH_BASE,
        sleeper: Callable[[float], Awaitable[None]] = _real_sleeper,
    ) -> None:
        self._http = http
        self._base = base_url.rstrip("/")
        self._sleep = sleeper

    # --- identidade ---
    async def me(self, access_token: str) -> dict:
        """`GET /me` — devolve a identidade do utilizador autenticado."""
        data = await self._get("/me", access_token)
        return {
            "id": data.get("id"),
            "displayName": data.get("displayName"),
            "userPrincipalName": data.get("userPrincipalName"),
        }

    # --- contactos: resolução de destinatários (read-only) ---
    async def search_people(
        self, access_token: str, query: str, *, top: int = 10
    ) -> list[dict]:
        """`GET /me/people?$search` — pessoas relevantes (ranking do Graph)."""
        data = await self._request(
            "GET", "/me/people", access_token,
            params={"$search": f'"{query}"', "$top": top},
        ) or {}
        out: list[dict] = []
        for p in data.get("value", []):
            email = None
            scored = p.get("scoredEmailAddresses") or []
            if scored:
                email = scored[0].get("address")
            if not email:
                emails = p.get("emailAddresses") or []
                if emails:
                    email = emails[0].get("address")
            out.append(
                {"display_name": p.get("displayName"), "email": email, "source": "people"}
            )
        return out

    async def search_contacts(
        self, access_token: str, query: str, *, top: int = 10
    ) -> list[dict]:
        """`GET /me/contacts?$search` — contactos pessoais do utilizador."""
        data = await self._request(
            "GET", "/me/contacts", access_token,
            params={"$search": f'"{query}"', "$top": top},
        ) or {}
        out: list[dict] = []
        for c in data.get("value", []):
            emails = c.get("emailAddresses") or []
            email = emails[0].get("address") if emails else None
            out.append(
                {"display_name": c.get("displayName"), "email": email, "source": "contacts"}
            )
        return out

    # --- email: leitura ---
    async def list_messages(
        self,
        access_token: str,
        *,
        search: str | None = None,
        filter_query: str | None = None,
        folder: str | None = None,
        top: int = 25,
        skip: int | None = None,
        select: str | None = None,
        orderby: str | None = None,
    ) -> dict:
        """`GET /me/messages` (ou de uma pasta). `$search` exige `ConsistencyLevel: eventual`."""
        path = (
            f"/me/mailFolders/{folder}/messages" if folder else "/me/messages"
        )
        params: dict[str, Any] = {"$top": top}
        headers: dict[str, str] = {}
        if search:
            params["$search"] = f'"{search}"'
            headers["ConsistencyLevel"] = "eventual"
        if filter_query:
            params["$filter"] = filter_query
        if skip is not None:
            params["$skip"] = skip
        if select:
            params["$select"] = select
        if orderby:
            params["$orderby"] = orderby
        data = await self._request(
            "GET", path, access_token, params=params, headers=headers
        ) or {}
        messages = [self._map_message_summary(m) for m in data.get("value", [])]
        return {"messages": messages, "next": data.get("@odata.nextLink")}

    async def list_messages_next(
        self, access_token: str, next_link: str, *, consistency: bool = False
    ) -> dict:
        """Segue um `@odata.nextLink` (URL absoluta) e devolve a página seguinte.

        A `nextLink` já traz a query (filtros + `$skiptoken`/`$skip`); quando a pesquisa
        original usou `$search`, é preciso repetir o cabeçalho `ConsistencyLevel: eventual`."""
        headers = {"ConsistencyLevel": "eventual"} if consistency else None
        data = await self._request(
            "GET", next_link, access_token, headers=headers
        ) or {}
        messages = [self._map_message_summary(m) for m in data.get("value", [])]
        return {"messages": messages, "next": data.get("@odata.nextLink")}

    async def get_message(
        self, access_token: str, message_id: str, *, select: str | None = None
    ) -> dict:
        """`GET /me/messages/{id}` — devolve a mensagem completa (com corpo)."""
        params = {"$select": select} if select else None
        data = await self._request(
            "GET", f"/me/messages/{message_id}", access_token, params=params
        ) or {}
        return {
            "id": data.get("id"),
            "subject": data.get("subject"),
            "from": self._addr(data.get("from")),
            "toRecipients": [self._addr(r) for r in data.get("toRecipients", [])],
            "ccRecipients": [self._addr(r) for r in data.get("ccRecipients", [])],
            "receivedDateTime": data.get("receivedDateTime"),
            "body": {
                "contentType": (data.get("body") or {}).get("contentType"),
                "content": (data.get("body") or {}).get("content"),
            },
            "hasAttachments": data.get("hasAttachments", False),
        }

    async def list_attachments(
        self, access_token: str, message_id: str
    ) -> list[dict]:
        """`GET /me/messages/{id}/attachments` — metadados dos anexos."""
        data = await self._request(
            "GET", f"/me/messages/{message_id}/attachments", access_token
        ) or {}
        return [
            {
                "id": a.get("id"),
                "name": a.get("name"),
                "contentType": a.get("contentType"),
                "size": a.get("size"),
                "isInline": a.get("isInline", False),
            }
            for a in data.get("value", [])
        ]

    async def get_attachment(
        self, access_token: str, message_id: str, attachment_id: str
    ) -> dict:
        """`GET /me/messages/{id}/attachments/{aid}` — inclui `contentBytes` (base64)."""
        data = await self._request(
            "GET",
            f"/me/messages/{message_id}/attachments/{attachment_id}",
            access_token,
        ) or {}
        return {
            "id": data.get("id"),
            "name": data.get("name"),
            "contentType": data.get("contentType"),
            "size": data.get("size"),
            "isInline": data.get("isInline", False),
            "contentBytes": data.get("contentBytes"),
        }

    # --- email: escrita ---
    async def send_mail(
        self, access_token: str, *, message: dict, save_to_sent_items: bool = True
    ) -> None:
        """`POST /me/sendMail` — envia uma mensagem montada no formato Graph."""
        await self._request(
            "POST",
            "/me/sendMail",
            access_token,
            json_body={
                "message": message,
                "saveToSentItems": save_to_sent_items,
            },
        )

    async def reply(
        self, access_token: str, message_id: str, *, comment: str, reply_all: bool = False
    ) -> None:
        """`POST /me/messages/{id}/reply` (ou `/replyAll`) — mantém a thread."""
        action = "replyAll" if reply_all else "reply"
        await self._request(
            "POST",
            f"/me/messages/{message_id}/{action}",
            access_token,
            json_body={"comment": comment},
        )

    async def forward(
        self,
        access_token: str,
        message_id: str,
        *,
        comment: str,
        to_recipients: list[str],
    ) -> None:
        """`POST /me/messages/{id}/forward` — reencaminha mantendo a thread."""
        await self._request(
            "POST",
            f"/me/messages/{message_id}/forward",
            access_token,
            json_body={
                "comment": comment,
                "toRecipients": recipients(to_recipients),
            },
        )

    async def list_folders(self, access_token: str) -> list[dict]:
        """`GET /me/mailFolders` — pastas de correio (id + nome)."""
        data = await self._request(
            "GET", "/me/mailFolders", access_token, params={"$top": 100}
        ) or {}
        return [
            {"id": f.get("id"), "displayName": f.get("displayName")}
            for f in data.get("value", [])
        ]

    async def move_message(
        self, access_token: str, message_id: str, *, destination_id: str
    ) -> dict:
        """`POST /me/messages/{id}/move` — devolve o novo recurso (na pasta destino)."""
        data = await self._request(
            "POST",
            f"/me/messages/{message_id}/move",
            access_token,
            json_body={"destinationId": destination_id},
        ) or {}
        return data

    async def delete_message(self, access_token: str, message_id: str) -> None:
        """`DELETE /me/messages/{id}` — soft delete (vai para Itens Eliminados).

        Mantido por compatibilidade; o soft delete das tools usa antes um `move` explícito
        para `deleteditems` (comportamento previsível e visível em Itens Eliminados)."""
        await self._request(
            "DELETE", f"/me/messages/{message_id}", access_token
        )

    async def permanent_delete(self, access_token: str, message_id: str) -> None:
        """`POST /me/messages/{id}/permanentDelete` — eliminação **permanente** (purges).

        Irrecuperável pelo utilizador (vai para a pasta `purges` do dumpster). Documentado em
        Graph v1.0; requer `Mail.ReadWrite`."""
        await self._request(
            "POST", f"/me/messages/{message_id}/permanentDelete", access_token
        )

    # --- email: anexos grandes (>3MB) via upload session ---
    async def create_draft(self, access_token: str, message: dict) -> dict:
        """`POST /me/messages` — cria um rascunho (base para anexos grandes)."""
        data = await self._request(
            "POST", "/me/messages", access_token, json_body=message
        ) or {}
        return data

    async def create_attachment_upload_session(
        self, access_token: str, message_id: str, *, attachment_item: dict
    ) -> dict:
        """`POST .../attachments/createUploadSession` — sessão de upload do anexo grande."""
        data = await self._request(
            "POST",
            f"/me/messages/{message_id}/attachments/createUploadSession",
            access_token,
            json_body={"AttachmentItem": attachment_item},
        ) or {}
        return data

    async def upload_attachment_bytes(
        self,
        upload_url: str,
        content_bytes: bytes,
        *,
        chunk_size: int = _UPLOAD_CHUNK_SIZE,
    ) -> None:
        """Carrega os bytes de um anexo grande para a `uploadUrl` da sessão, em chunks.

        A `uploadUrl` é pré-autenticada (SAS-like) — NÃO se envia o `Bearer`; é um PUT a uma
        URL absoluta. Cada chunk leva `Content-Range: bytes {ini}-{fim}/{total}`. O Graph
        aceita chunks múltiplos de 320 KiB; o último pode ser menor e devolve 201/200.
        """
        total = len(content_bytes)
        if total == 0:
            return
        start = 0
        while start < total:
            end = min(start + chunk_size, total) - 1
            chunk = content_bytes[start : end + 1]
            headers = {
                "Content-Length": str(len(chunk)),
                "Content-Range": f"bytes {start}-{end}/{total}",
            }
            resp = await self._http.put(upload_url, content=chunk, headers=headers)
            if resp.status_code >= 400:
                raise GraphError(
                    f"Upload de anexo falhou ({resp.status_code}): {resp.text[:200]}"
                )
            start = end + 1

    async def send_draft(self, access_token: str, message_id: str) -> None:
        """`POST /me/messages/{id}/send` — envia um rascunho previamente criado."""
        await self._request(
            "POST", f"/me/messages/{message_id}/send", access_token
        )

    # --- calendário: fuso do mailbox (D1) ---
    async def get_mailbox_timezone(self, access_token: str) -> str | None:
        """`GET /me/mailboxSettings` -> devolve `timeZone` (ex.: 'GMT Standard Time').

        None se ausente. Lido uma vez por pedido e reutilizado em leituras/escritas. O valor
        vem TAL E QUAL do mailbox (Windows ou IANA) — o Graph aceita ambos no header
        `Prefer`/`timeZone`; não convertemos (R2)."""
        data = await self._request(
            "GET", "/me/mailboxSettings", access_token
        ) or {}
        return data.get("timeZone")

    # --- calendário: leitura (D5 auto-pagina; D1 header Prefer) ---
    @staticmethod
    def _prefer_tz_headers(prefer_timezone: str | None) -> dict | None:
        """Header `Prefer: outlook.timezone="<tz>"` quando há fuso; senão None."""
        if not prefer_timezone:
            return None
        return {"Prefer": f'outlook.timezone="{prefer_timezone}"'}

    async def list_calendar_view(
        self,
        access_token: str,
        *,
        start: str,
        end: str,
        top: int = 50,
        prefer_timezone: str | None = None,
    ) -> dict:
        """`GET /me/calendarView?startDateTime=&endDateTime=` — expande ocorrências de séries
        no intervalo (D4: ler recorrências já expandidas). `$orderby=start/dateTime`.
        Header `Prefer: outlook.timezone="<tz>"` quando `prefer_timezone`.
        Devolve {"events": [...summary...], "next": data.get("@odata.nextLink")}."""
        params = {
            "startDateTime": start,
            "endDateTime": end,
            "$orderby": "start/dateTime",
            "$top": top,
        }
        data = await self._request(
            "GET", "/me/calendarView", access_token,
            params=params, headers=self._prefer_tz_headers(prefer_timezone),
        ) or {}
        events = [self._map_event_summary(e) for e in data.get("value", [])]
        return {"events": events, "next": data.get("@odata.nextLink")}

    async def list_calendar_view_next(
        self, access_token: str, next_link: str, *, prefer_timezone: str | None = None
    ) -> dict:
        """Segue um `@odata.nextLink` absoluto do calendarView. Repete o header `Prefer`
        (o fuso não viaja na nextLink). Devolve {"events": [...], "next": ...}."""
        data = await self._request(
            "GET", next_link, access_token,
            headers=self._prefer_tz_headers(prefer_timezone),
        ) or {}
        events = [self._map_event_summary(e) for e in data.get("value", [])]
        return {"events": events, "next": data.get("@odata.nextLink")}

    async def get_event(
        self, access_token: str, event_id: str, *, prefer_timezone: str | None = None
    ) -> dict:
        """`GET /me/events/{id}` — evento completo (com corpo). Header `Prefer` quando dado.
        Devolve _map_event_detail(data)."""
        data = await self._request(
            "GET", f"/me/events/{event_id}", access_token,
            headers=self._prefer_tz_headers(prefer_timezone),
        ) or {}
        return self._map_event_detail(data)

    # --- calendário: disponibilidade (D2) ---
    async def get_schedule(
        self,
        access_token: str,
        *,
        schedules: list[str],
        start: str,
        end: str,
        interval_minutes: int = 30,
        prefer_timezone: str | None = None,
    ) -> list[dict]:
        """`POST /me/calendar/getSchedule` — free/busy do próprio + participantes (D2).

        Devolve, por schedule, {email, availabilityView, scheduleItems:[{status,start,end}]}.
        Header `Prefer` quando dado."""
        tz = prefer_timezone or "UTC"
        body = {
            "schedules": schedules,
            "startTime": {"dateTime": start, "timeZone": tz},
            "endTime": {"dateTime": end, "timeZone": tz},
            "availabilityViewInterval": interval_minutes,
        }
        data = await self._request(
            "POST", "/me/calendar/getSchedule", access_token,
            json_body=body, headers=self._prefer_tz_headers(prefer_timezone),
        ) or {}
        out: list[dict] = []
        for s in data.get("value", []):
            items = [
                {
                    "status": it.get("status"),
                    "start": (it.get("start") or {}).get("dateTime"),
                    "end": (it.get("end") or {}).get("dateTime"),
                }
                for it in s.get("scheduleItems", [])
            ]
            out.append(
                {
                    "email": s.get("scheduleId"),
                    "availabilityView": s.get("availabilityView"),
                    "scheduleItems": items,
                }
            )
        return out

    # --- calendário: escrita ---
    async def create_event(self, access_token: str, *, event: dict) -> dict:
        """`POST /me/events` — cria o evento (objeto Graph montado pela tool). Devolve o
        recurso criado (mapeado por _map_event_detail; expõe pelo menos id e webLink)."""
        data = await self._request(
            "POST", "/me/events", access_token, json_body=event
        ) or {}
        return self._map_event_detail(data)

    async def update_event(
        self, access_token: str, event_id: str, *, changes: dict
    ) -> dict:
        """`PATCH /me/events/{id}` — aplica só os campos alterados. Devolve o atualizado."""
        data = await self._request(
            "PATCH", f"/me/events/{event_id}", access_token, json_body=changes
        ) or {}
        return self._map_event_detail(data)

    async def cancel_event(
        self, access_token: str, event_id: str, *, comment: str = ""
    ) -> None:
        """`POST /me/events/{id}/cancel` — cancela e notifica os participantes (organizador).
        202/204 -> None. (Só o organizador pode cancelar — ver D7/US-2.5.)"""
        await self._request(
            "POST", f"/me/events/{event_id}/cancel", access_token,
            json_body={"comment": comment},
        )

    async def respond_event(
        self,
        access_token: str,
        event_id: str,
        *,
        response: str,
        comment: str = "",
        send_response: bool = True,
    ) -> None:
        """`POST /me/events/{id}/{response}` — responde ao convite (accept|decline|
        tentativelyAccept). 202/204 -> None."""
        await self._request(
            "POST", f"/me/events/{event_id}/{response}", access_token,
            json_body={"comment": comment, "sendResponse": send_response},
        )

    # --- núcleo HTTP ---
    async def _get(self, path: str, access_token: str) -> dict:
        """Compatibilidade Fase 0: GET simples que delega no `_request`."""
        return await self._request("GET", path, access_token) or {}

    async def _request(
        self,
        method: str,
        path: str,
        access_token: str,
        *,
        params: dict | None = None,
        json_body: dict | None = None,
        headers: dict | None = None,
    ) -> dict | None:
        """Pedido genérico ao Graph com a política de retry/erros partilhada.

        Retry em 429 (respeita `Retry-After`); 401/403 -> `UpstreamAuthError`; >=400 ->
        `GraphError`. Respostas 202/204 (ou sem corpo) devolvem `None`.
        """
        # `path` pode já ser uma URL absoluta (ex.: um `@odata.nextLink` de paginação).
        url = path if path.startswith(("http://", "https://")) else f"{self._base}{path}"
        req_headers = {"Authorization": f"Bearer {access_token}"}
        if headers:
            req_headers.update(headers)
        attempts = 0
        while True:
            resp = await self._http.request(
                method, url, headers=req_headers, params=params, json=json_body
            )
            if resp.status_code == 429 and attempts < _MAX_RETRIES_429:
                retry_after = float(resp.headers.get("Retry-After", "1"))
                await self._sleep(retry_after)
                attempts += 1
                continue
            if resp.status_code in (401, 403):
                raise UpstreamAuthError(
                    f"Graph rejeitou o token ({resp.status_code})."
                )
            if resp.status_code >= 400:
                raise GraphError(
                    f"Graph devolveu {resp.status_code}: {resp.text[:200]}"
                )
            if resp.status_code in (202, 204) or not resp.content:
                return None
            return resp.json()

    # --- mapeamentos ---
    @staticmethod
    def _addr(recipient: dict | None) -> str | None:
        """Extrai o endereço de um objeto recipient do Graph (`{emailAddress:{address}}`)."""
        if not recipient:
            return None
        return (recipient.get("emailAddress") or {}).get("address")

    @classmethod
    def _map_message_summary(cls, m: dict) -> dict:
        return {
            "id": m.get("id"),
            "subject": m.get("subject"),
            "from": cls._addr(m.get("from")),
            "receivedDateTime": m.get("receivedDateTime"),
            "bodyPreview": m.get("bodyPreview"),
            "hasAttachments": m.get("hasAttachments", False),
            "isRead": m.get("isRead"),
        }

    # --- mapeamentos: calendário ---
    @staticmethod
    def _is_recurring(e: dict) -> bool:
        """True se a ocorrência pertence a/ é uma série recorrente."""
        if e.get("seriesMasterId"):
            return True
        return e.get("type") in ("occurrence", "exception", "seriesMaster")

    @classmethod
    def _map_event_summary(cls, e: dict) -> dict:
        return {
            "id": e.get("id"),
            "subject": e.get("subject"),
            "start": {
                "dateTime": (e.get("start") or {}).get("dateTime"),
                "timeZone": (e.get("start") or {}).get("timeZone"),
            },
            "end": {
                "dateTime": (e.get("end") or {}).get("dateTime"),
                "timeZone": (e.get("end") or {}).get("timeZone"),
            },
            "location": (e.get("location") or {}).get("displayName"),
            "organizer": ((e.get("organizer") or {}).get("emailAddress") or {}).get(
                "address"
            ),
            "isOnlineMeeting": e.get("isOnlineMeeting", False),
            "joinUrl": (e.get("onlineMeeting") or {}).get("joinUrl"),
            "isRecurring": cls._is_recurring(e),
            "seriesMasterId": e.get("seriesMasterId"),
            "isAllDay": e.get("isAllDay", False),
            "bodyPreview": e.get("bodyPreview"),
        }

    @classmethod
    def _map_event_detail(cls, e: dict) -> dict:
        detail = cls._map_event_summary(e)
        detail.update(
            {
                "attendees": [
                    {
                        "email": ((a.get("emailAddress") or {}).get("address")),
                        "name": ((a.get("emailAddress") or {}).get("name")),
                        "type": a.get("type"),
                        "responseStatus": ((a.get("status") or {}).get("response")),
                    }
                    for a in e.get("attendees", [])
                ],
                "responseStatus": (e.get("responseStatus") or {}).get("response"),
                "body": {
                    "contentType": (e.get("body") or {}).get("contentType"),
                    "content": (e.get("body") or {}).get("content"),
                },
                "webLink": e.get("webLink"),
                "type": e.get("type"),
            }
        )
        return detail
