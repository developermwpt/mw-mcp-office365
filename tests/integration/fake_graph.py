"""FakeGraphClient para os testes end-to-end das tools de email.

Implementa a mesma interface assíncrona do `GraphClient` real, mas SEM rede: cada método
regista a chamada (em `self.calls`) e devolve uma resposta programada. Contar chamadas é o
que prova as duas invariantes-chave da Fase 1:

- "prepare NÃO toca no Graph" (nenhuma chamada de escrita até ao confirm);
- idempotência (um segundo confirm não duplica o envio/eliminação).
"""

from __future__ import annotations

from typing import Any

from mcp_o365.auth.errors import UpstreamAuthError


class FakeGraphClient:
    """Fake do GraphClient: regista chamadas e devolve respostas programáveis."""

    def __init__(
        self,
        *,
        messages: dict | None = None,
        message: dict | None = None,
        attachments: list[dict] | None = None,
        attachment: dict | None = None,
        folders: list[dict] | None = None,
        moved: dict | None = None,
        draft: dict | None = None,
        people: list[dict] | None = None,
        contacts: list[dict] | None = None,
        next_pages: list[dict] | None = None,
        auth_fail: dict[str, int] | None = None,
        # --- calendário (Fase 2) ---
        me: dict | None = None,
        mailbox_timezone: str | None = None,
        events: dict | None = None,
        next_event_pages: list[dict] | None = None,
        event: dict | None = None,
        schedule: list[dict] | None = None,
        # --- Teams (Fase 3) ---
        chats: dict | None = None,
        next_chat_pages: list[dict] | None = None,
        chat: dict | None = None,
        chat_messages: dict | None = None,
        next_message_pages: list[dict] | None = None,
        created_chat: dict | None = None,
        sent_message: dict | None = None,
        # --- agendamento de envio (US-1.9/1.10/1.11) ---
        deferred_drafts: dict | None = None,
    ) -> None:
        self.calls: list[tuple[str, tuple, dict]] = []
        self._messages = messages or {"messages": [], "next": None}
        # Páginas seguintes devolvidas por `list_messages_next` (consumidas por ordem).
        self._next_pages = list(next_pages or [])
        self._next_idx = 0
        self._message = message or {}
        self._attachments = attachments or []
        self._attachment = attachment or {}
        self._folders = folders or []
        self._moved = moved or {"id": "msg-novo"}
        self._draft = draft or {"id": "draft-1"}
        self._people = people if people is not None else []
        self._contacts = contacts if contacts is not None else []
        # Nº de vezes que cada método deve simular um 401/403 do Graph antes de ter sucesso.
        self._auth_fail = dict(auth_fail or {})
        # --- calendário (Fase 2) ---
        self._me = me or {"userPrincipalName": "subj@example.com"}
        self._mailbox_timezone = mailbox_timezone
        self._events = events or {"events": [], "next": None}
        # Páginas seguintes devolvidas por `list_calendar_view_next` (por ordem).
        self._next_event_pages = list(next_event_pages or [])
        self._next_event_idx = 0
        self._event = event or {"id": "evt-1", "webLink": "https://web/evt-1"}
        self._schedule = schedule if schedule is not None else []
        # --- Teams (Fase 3) ---
        self._chats = chats or {"chats": [], "next": None}
        self._next_chat_pages = list(next_chat_pages or [])
        self._next_chat_idx = 0
        self._chat = chat or {}
        self._chat_messages = chat_messages or {"messages": [], "next": None}
        self._next_message_pages = list(next_message_pages or [])
        self._next_message_idx = 0
        self._created_chat = created_chat or {"id": "chat-novo", "chat_type": "oneOnOne"}
        self._sent_message = sent_message or {"id": "msg-1", "created": "2026-06-06T10:00:00Z"}
        # --- agendamento de envio (US-1.9/1.10/1.11) ---
        # Já no formato mapeado pela tool: {"drafts": [{"id","subject","to":[...],
        # "deferred_send_at":"...Z"}], "next": None}.
        self._deferred_drafts = deferred_drafts or {"drafts": [], "next": None}

    def _record(self, name: str, *args: Any, **kwargs: Any) -> None:
        self.calls.append((name, args, kwargs))
        remaining = self._auth_fail.get(name, 0)
        if remaining > 0:
            self._auth_fail[name] = remaining - 1
            raise UpstreamAuthError(f"Graph rejeitou o token (simulado em {name}).")

    def count(self, name: str) -> int:
        return sum(1 for c in self.calls if c[0] == name)

    # --- leitura ---
    async def list_messages(self, access_token, **kwargs) -> dict:
        self._record("list_messages", access_token, **kwargs)
        return self._messages

    async def list_messages_next(self, access_token, next_link, *, consistency=False) -> dict:
        self._record("list_messages_next", access_token, next_link, consistency=consistency)
        if self._next_idx < len(self._next_pages):
            page = self._next_pages[self._next_idx]
            self._next_idx += 1
            return {"messages": page.get("messages", []), "next": page.get("next")}
        return {"messages": [], "next": None}

    async def get_message(self, access_token, message_id, **kwargs) -> dict:
        self._record("get_message", access_token, message_id, **kwargs)
        return self._message

    async def list_attachments(self, access_token, message_id) -> list[dict]:
        self._record("list_attachments", access_token, message_id)
        return self._attachments

    async def get_attachment(self, access_token, message_id, attachment_id) -> dict:
        self._record("get_attachment", access_token, message_id, attachment_id)
        return self._attachment

    # --- escrita ---
    async def send_mail(self, access_token, *, message, save_to_sent_items=True) -> None:
        self._record("send_mail", access_token, message=message,
                     save_to_sent_items=save_to_sent_items)

    async def reply(self, access_token, message_id, *, comment, reply_all=False) -> None:
        self._record("reply", access_token, message_id, comment=comment,
                     reply_all=reply_all)

    async def forward(self, access_token, message_id, *, comment, to_recipients) -> None:
        self._record("forward", access_token, message_id, comment=comment,
                     to_recipients=to_recipients)

    async def list_folders(self, access_token) -> list[dict]:
        self._record("list_folders", access_token)
        return self._folders

    async def move_message(self, access_token, message_id, *, destination_id) -> dict:
        self._record("move_message", access_token, message_id,
                     destination_id=destination_id)
        return self._moved

    async def delete_message(self, access_token, message_id) -> None:
        self._record("delete_message", access_token, message_id)

    async def permanent_delete(self, access_token, message_id) -> None:
        self._record("permanent_delete", access_token, message_id)

    # --- identidade ---
    async def me(self, access_token) -> dict:
        self._record("me", access_token)
        return self._me

    # --- calendário: fuso + leitura (Fase 2) ---
    async def get_mailbox_timezone(self, access_token) -> str | None:
        self._record("get_mailbox_timezone", access_token)
        return self._mailbox_timezone

    async def list_calendar_view(self, access_token, **kwargs) -> dict:
        self._record("list_calendar_view", access_token, **kwargs)
        return self._events

    async def list_calendar_view_next(
        self, access_token, next_link, *, prefer_timezone=None
    ) -> dict:
        self._record("list_calendar_view_next", access_token, next_link,
                     prefer_timezone=prefer_timezone)
        if self._next_event_idx < len(self._next_event_pages):
            page = self._next_event_pages[self._next_event_idx]
            self._next_event_idx += 1
            return {"events": page.get("events", []), "next": page.get("next")}
        return {"events": [], "next": None}

    async def get_event(self, access_token, event_id, *, prefer_timezone=None) -> dict:
        self._record("get_event", access_token, event_id, prefer_timezone=prefer_timezone)
        return self._event

    # --- calendário: disponibilidade (Fase 2) ---
    async def get_schedule(
        self, access_token, *, schedules, start, end,
        interval_minutes=30, prefer_timezone=None,
    ) -> list[dict]:
        self._record("get_schedule", access_token, schedules=schedules, start=start,
                     end=end, interval_minutes=interval_minutes,
                     prefer_timezone=prefer_timezone)
        return self._schedule

    # --- calendário: escrita (Fase 2) ---
    async def create_event(self, access_token, *, event) -> dict:
        self._record("create_event", access_token, event=event)
        return self._event

    async def update_event(self, access_token, event_id, *, changes) -> dict:
        self._record("update_event", access_token, event_id, changes=changes)
        return self._event

    async def cancel_event(self, access_token, event_id, *, comment="") -> None:
        self._record("cancel_event", access_token, event_id, comment=comment)

    async def respond_event(
        self, access_token, event_id, *, response, comment="", send_response=True
    ) -> None:
        self._record("respond_event", access_token, event_id, response=response,
                     comment=comment, send_response=send_response)

    # --- Teams: leitura (Fase 3) ---
    async def list_chats(self, access_token, *, top=50) -> dict:
        self._record("list_chats", access_token, top=top)
        return self._chats

    async def list_chats_next(self, access_token, next_link) -> dict:
        self._record("list_chats_next", access_token, next_link)
        if self._next_chat_idx < len(self._next_chat_pages):
            page = self._next_chat_pages[self._next_chat_idx]
            self._next_chat_idx += 1
            return {"chats": page.get("chats", []), "next": page.get("next")}
        return {"chats": [], "next": None}

    async def get_chat(self, access_token, chat_id) -> dict:
        self._record("get_chat", access_token, chat_id)
        return self._chat

    async def list_chat_messages(self, access_token, chat_id, *, top=25) -> dict:
        self._record("list_chat_messages", access_token, chat_id, top=top)
        return self._chat_messages

    async def list_chat_messages_next(self, access_token, next_link) -> dict:
        self._record("list_chat_messages_next", access_token, next_link)
        if self._next_message_idx < len(self._next_message_pages):
            page = self._next_message_pages[self._next_message_idx]
            self._next_message_idx += 1
            return {"messages": page.get("messages", []), "next": page.get("next")}
        return {"messages": [], "next": None}

    # --- Teams: escrita (Fase 3) ---
    async def create_one_on_one_chat(self, access_token, *, member_emails) -> dict:
        self._record("create_one_on_one_chat", access_token, member_emails=member_emails)
        return self._created_chat

    async def send_chat_message(
        self, access_token, chat_id, *, content, content_type="text"
    ) -> dict:
        self._record("send_chat_message", access_token, chat_id, content=content,
                     content_type=content_type)
        return self._sent_message

    # --- contactos ---
    async def search_people(self, access_token, query, *, top=10) -> list[dict]:
        self._record("search_people", access_token, query, top=top)
        return self._people

    async def search_contacts(self, access_token, query, *, top=10) -> list[dict]:
        self._record("search_contacts", access_token, query, top=top)
        return self._contacts

    # --- anexos grandes (upload session) ---
    async def create_draft(self, access_token, message) -> dict:
        self._record("create_draft", access_token, message)
        return self._draft

    async def create_attachment_upload_session(
        self, access_token, message_id, *, attachment_item
    ) -> dict:
        self._record("create_attachment_upload_session", access_token, message_id,
                     attachment_item=attachment_item)
        return {"uploadUrl": "https://upload.example/x"}

    async def upload_attachment_bytes(
        self, upload_url, content_bytes, **kwargs
    ) -> None:
        self._record("upload_attachment_bytes", upload_url,
                     content_bytes=content_bytes, **kwargs)

    async def send_draft(self, access_token, message_id) -> None:
        self._record("send_draft", access_token, message_id)

    # --- agendamento de envio (US-1.10): listar rascunhos diferidos ---
    async def list_deferred_drafts(self, access_token, *, prop_id, top=50) -> dict:
        self._record("list_deferred_drafts", access_token, prop_id=prop_id, top=top)
        return self._deferred_drafts
