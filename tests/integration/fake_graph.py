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
        auth_fail: dict[str, int] | None = None,
    ) -> None:
        self.calls: list[tuple[str, tuple, dict]] = []
        self._messages = messages or {"messages": [], "next": None}
        self._message = message or {}
        self._attachments = attachments or []
        self._attachment = attachment or {}
        self._folders = folders or []
        self._moved = moved or {"id": "msg-novo"}
        self._draft = draft or {"id": "draft-1"}
        # Nº de vezes que cada método deve simular um 401/403 do Graph antes de ter sucesso.
        self._auth_fail = dict(auth_fail or {})

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
