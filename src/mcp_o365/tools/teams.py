"""Fase 3 — Tools de Teams (chats 1:1 e de grupo; US-3.1 a US-3.5).

Funções `run_teams_*` standalone, independentes do transporte MCP, testáveis com Graph e
Entra mockados. Seguem à risca os padrões da Fase 1/2 (`tools/calendar.py`, `tools/email.py`):
resolução de token via `resolve_access_token`/`call_graph`, reautenticação graciosa
(`reauth_response`), o par prepare/confirm do `ApprovalEngine` para escritas, auditoria
só-metadados (`log_audit`) e sanitização de conteúdo não-confiável (`sanitize_html`).

Decisões fechadas implementadas (D1–D11, ver docs/fase-3/plano-implementacao.md):
- D1 — criar chat 1:1 quando não existe = ESCRITA confirmada (prepare/confirm).
- D2 — filtro de listagem CLIENT-SIDE (tópico de grupo OU nome/email de membro).
- D3 — obtenção do chat 1:1: procurar primeiro (oneOnOne cujo único outro membro == email),
  só criar depois (via prepare/confirm). Match por email, case-insensitive, excluindo o
  próprio; membro sem email -> sem match -> cria (idempotente no Graph) — comportamento
  esperado (achado A5 da revisão do coordenador).
- D4 — limite de mensagens lidas: default 25, teto 50.
- D5 — histórico = N mais recentes + `has_more`/`next_link` (NÃO auto-pagina, NÃO pergunta).
- D6 — formato de envio: default `text`; `html` só a pedido explícito.
- D8 — mensagens de sistema marcadas (`is_system`), nunca interpretadas como acionáveis.
- D9 — resolução por nome a montante (`resolve_recipient` + confirmação humana antes).
- D10 — limite de tamanho do corpo (`_MAX_BODY_CHARS`); acima -> erro orientador.

Correções da revisão do coordenador incorporadas:
- A1 — a auditoria `teams.send` NÃO leva `subject_hash` em `extra` (o `log_audit` já emite o
  `subject_hash` da IDENTIDADE de topo; injetar um segundo sobrescreveria-o). `extra` de
  `teams.send` = `{chat_type, body_type}`. Por isso NÃO importamos `subject_hash` aqui.
- A2 — o resumo do `teams_send_message_prepare` é montado a partir de `get_chat(chat_id)`
  (leitura pontual, robusta a >50 chats), não de um match sobre `list_chats`.
"""

from __future__ import annotations

import logging
from collections.abc import Callable
from datetime import datetime, timezone

from ..approval.engine import ApprovalEngine
from ..approval.errors import ConfirmationExpired, ConfirmationNotFound
from ..auth.errors import ReauthRequired, UpstreamAuthError
from ..auth.plane_b import PlaneB
from ..graph.client import GraphClient
from ..graph.sanitize import sanitize_html
from ..identity.mapping import IdentityMapping
from ..observability.audit import log_audit
from ..storage.token_store import TokenStore
from ._session import call_graph, reauth_response, resolve_access_token
from .calendar import _own_email
from .email import _domains

logger = logging.getLogger("mcp_o365.tools.teams")
audit_logger = logging.getLogger("mcp_o365.audit")

_MAX_MESSAGES_PER_CALL = 50        # teto de mensagens por leitura (D4)
_DEFAULT_MESSAGES = 25             # default de leitura (D4), igual ao email
_MAX_BODY_CHARS = 28000            # teto de tamanho da mensagem (D10)
_VALID_BODY_TYPES = {"text", "html"}   # D6
_MAX_LIST_FETCH = 200              # teto da paginação acessória da listagem (D2 client-side)


def _utcnow() -> datetime:
    return datetime.now(timezone.utc)


def _sanitize_preview(chat: dict) -> dict:
    """Sanitiza o `last_message_preview` de um chat resumido (conteúdo não-confiável).

    O preview vem CRU do client (R6: pode ser None); aqui passa por `sanitize_html` antes de
    chegar ao modelo, como no `bodyPreview` do email/calendário."""
    preview = chat.get("last_message_preview")
    if preview:
        return {**chat, "last_message_preview": sanitize_html(preview)}
    return chat


def _sanitize_message(message: dict) -> dict:
    """Sanitiza o corpo HTML de uma mensagem (conteúdo não-confiável).

    O `_map_chat_message` devolve o corpo CRU; quando `contentType == 'html'`, sanitiza-se
    aqui (a fronteira de sanitização fica na tool). Mantém `is_system`/metadados intactos."""
    body = message.get("body") or {}
    if (body.get("contentType") == "html") and body.get("content"):
        return {
            **message,
            "body": {**body, "content": sanitize_html(body["content"])},
        }
    return message


def _chat_type_label(chat_type: str | None) -> str:
    """Rótulo PT do tipo de chat para o resumo de confirmação."""
    return "1:1" if chat_type == "oneOnOne" else "de grupo"


# ============================ LEITURA (sem aprovação) ============================


async def run_teams_list_chats(
    subject: str | None,
    *,
    mapping: IdentityMapping,
    plane_b: PlaneB,
    graph_client: GraphClient,
    store: TokenStore,
    filter_text: str | None = None,
    top: int = 50,
    account_id: str | None = None,
    clock: Callable[[], datetime] = _utcnow,
) -> dict:
    """US-3.1 — Lista chats 1:1 e de grupo com metadados (read-only).

    Endpoint `GET /me/chats?$expand=members,lastMessagePreview&$top={top}` via `list_chats`.
    Se `filter_text` e houver `next`, pagina via `list_chats_next` até satisfazer o filtro ou
    atingir `_MAX_LIST_FETCH` (D2 client-side). Filtra em memória (case-insensitive substring)
    por `topic` OU nome/email de qualquer membro. Preview sanitizado (não-confiável)."""
    try:
        _, first = await call_graph(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            op=lambda token: graph_client.list_chats(token, top=top),
            account_id=account_id, clock=clock,
        )
        all_chats = list(first["chats"])
        next_link = first["next"]
        truncated = False
        # Só vale a pena paginar quando há filtro (D2): a 1ª página chega para listar.
        while next_link and filter_text:
            if len(all_chats) >= _MAX_LIST_FETCH:
                truncated = True
                break

            async def fetch_more(token: str, _link: str = next_link) -> dict:
                return await graph_client.list_chats_next(token, _link)

            _, page = await call_graph(
                subject, mapping=mapping, plane_b=plane_b, store=store,
                op=fetch_more, account_id=account_id, clock=clock,
            )
            all_chats.extend(page["chats"])
            next_link = page["next"]
    except ReauthRequired as exc:
        return reauth_response(str(exc))

    if filter_text:
        needle = filter_text.casefold()

        def _matches(chat: dict) -> bool:
            if (chat.get("topic") or "").casefold().find(needle) >= 0:
                return True
            for m in chat.get("members") or []:
                if (m.get("name") or "").casefold().find(needle) >= 0:
                    return True
                if (m.get("email") or "").casefold().find(needle) >= 0:
                    return True
            return False

        all_chats = [c for c in all_chats if _matches(c)]

    all_chats = [_sanitize_preview(c) for c in all_chats]
    # Ordena defensivamente por last_updated desc (caso o Graph não ordene — A10).
    all_chats.sort(key=lambda c: c.get("last_updated") or "", reverse=True)

    result = {
        "status": "ok",
        "chats": all_chats,
        "count": len(all_chats),
        "has_more": bool(next_link),
        "content_is_untrusted": True,
    }
    if truncated:
        # A7 — sinaliza truncagem em vez de truncar em silêncio.
        result["truncated_at"] = _MAX_LIST_FETCH
    return result


async def run_teams_read_messages(
    subject: str | None,
    *,
    mapping: IdentityMapping,
    plane_b: PlaneB,
    graph_client: GraphClient,
    store: TokenStore,
    chat_id: str,
    top: int = _DEFAULT_MESSAGES,
    page_token: str | None = None,
    account_id: str | None = None,
    clock: Callable[[], datetime] = _utcnow,
) -> dict:
    """US-3.2 — Lê as N mensagens MAIS RECENTES de um chat (read-only).

    `top` é fixado em `[1, _MAX_MESSAGES_PER_CALL]` (D4). Sem `page_token` ->
    `list_chat_messages`; com `page_token` -> `list_chat_messages_next` (mensagens mais
    antigas, a pedido — D5). NÃO auto-pagina. Corpo HTML sanitizado; `is_system` já vem
    marcado (D8); resposta com `content_is_untrusted`."""
    if not chat_id:
        return {"status": "error", "message": "Indique o chat_id do chat a ler."}

    top = min(max(top, 1), _MAX_MESSAGES_PER_CALL)

    try:
        if page_token:
            async def fetch(token: str) -> dict:
                return await graph_client.list_chat_messages_next(token, page_token)
        else:
            async def fetch(token: str) -> dict:
                return await graph_client.list_chat_messages(token, chat_id, top=top)

        _, data = await call_graph(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            op=fetch, account_id=account_id, clock=clock,
        )
    except ReauthRequired as exc:
        return reauth_response(str(exc))

    messages = [_sanitize_message(m) for m in data["messages"]]
    next_link = data["next"]
    return {
        "status": "ok",
        "chat_id": chat_id,
        "messages": messages,
        "count": len(messages),
        "has_more": bool(next_link),
        "next_link": next_link,
        "content_is_untrusted": True,
    }


# ====================== ESCRITA (prepare / confirm) ======================


async def run_teams_send_message_prepare(
    subject: str | None,
    *,
    mapping: IdentityMapping,
    plane_b: PlaneB,
    graph_client: GraphClient,
    store: TokenStore,
    approval: ApprovalEngine,
    chat_id: str,
    body: str,
    body_type: str = "text",
    account_id: str | None = None,
    clock: Callable[[], datetime] = _utcnow,
) -> dict:
    """US-3.3/US-3.5 — Prepara o envio de uma mensagem para um chat EXISTENTE (NÃO envia).

    Validação antes de qualquer leitura/escrita (D6/D10). Leitura acessória best-effort via
    `get_chat(chat_id)` (A2) para montar o resumo (tipo de chat + nº participantes + domínios);
    se falhar por motivo não-auth, degrada graciosamente (resumo sem detalhes), nunca escreve.
    Responder numa conversa = enviar no mesmo `chat_id` (D7: não há thread em chats)."""
    if not chat_id or not body:
        return {"status": "error", "message": "Indique o chat_id e o corpo da mensagem (body)."}
    if body_type not in _VALID_BODY_TYPES:
        return {"status": "error", "message": "Formato inválido. Use 'text' ou 'html'."}
    if len(body) > _MAX_BODY_CHARS:
        return {
            "status": "error",
            "message": (
                f"Mensagem demasiado longa ({len(body)} caracteres; máximo "
                f"{_MAX_BODY_CHARS}). Divida em partes."
            ),
        }

    try:
        account, token = await resolve_access_token(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            account_id=account_id, clock=clock,
        )
    except ReauthRequired as exc:
        return reauth_response(str(exc))

    # Leitura acessória (A2): get_chat para o resumo. Best-effort por desenho — NÃO passa por
    # `call_graph` para que uma falha aqui (scope/permissão em falta -> 401/403) não force
    # refresh nem marque a conta como expirada (mesmo cuidado do `_resolve_tz` da Fase 2). Em
    # falha, degrada o resumo (sem detalhes do chat) mas ainda emite o token — nunca escreve.
    chat: dict = {}
    try:
        chat = await graph_client.get_chat(token, chat_id) or {}
    except (UpstreamAuthError, ReauthRequired):
        chat = {}

    members = chat.get("members") or []
    chat_type = chat.get("chat_type")
    n = len(members)
    member_emails = [m["email"] for m in members if m.get("email")]
    if chat_type:
        summary = (
            f"Enviar mensagem no chat {_chat_type_label(chat_type)} com {n} participante(s) "
            f"(domínios: {', '.join(_domains(member_emails)) or 'n/d'}) [formato: {body_type}]."
        )
    else:
        # get_chat degradou (best-effort): resumo sem detalhes do chat, mas ainda emite token.
        summary = (
            f"Enviar mensagem no chat {chat_id} [formato: {body_type}]. "
            "(Detalhes do chat indisponíveis.)"
        )

    prepared = approval.prepare(
        subject=subject,
        account_id=account.account_id,
        operation="teams.send",
        payload={
            "chat_id": chat_id,
            "content": body,
            "content_type": body_type,
            "chat_type": chat_type,
            "recipients_count": n,
        },
        summary=summary,
    )
    prepared["recipients_count"] = n
    return prepared


async def run_teams_send_message_confirm(
    subject: str | None,
    *,
    mapping: IdentityMapping,
    plane_b: PlaneB,
    graph_client: GraphClient,
    store: TokenStore,
    approval: ApprovalEngine,
    confirmation_token: str,
    account_id: str | None = None,
    clock: Callable[[], datetime] = _utcnow,
) -> dict:
    """US-3.3/US-3.5 — Confirma e envia a mensagem preparada (token fresco); audita
    `teams.send` (só-metadados; A1: `extra={chat_type, body_type}`, SEM subject_hash)."""
    async def executor(operation: str, payload: dict) -> dict:
        account, created = await call_graph(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            op=lambda token: graph_client.send_chat_message(
                token, payload["chat_id"],
                content=payload["content"], content_type=payload["content_type"],
            ),
            account_id=account_id, clock=clock,
        )
        log_audit(
            audit_logger, action="teams.send", subject=subject,
            account_id=account.account_id, target=payload["chat_id"],
            outcome="success", recipients_count=payload.get("recipients_count"),
            extra={
                "chat_type": payload.get("chat_type"),
                "body_type": payload.get("content_type"),
            },
        )
        return {
            "operation": operation,
            "chat_id": payload["chat_id"],
            "message_id": (created or {}).get("id"),
            "message": "Mensagem enviada.",
        }

    return await _confirm(approval, subject=subject, token=confirmation_token,
                          executor=executor)


async def run_teams_get_or_create_one_on_one_chat_prepare(
    subject: str | None,
    *,
    mapping: IdentityMapping,
    plane_b: PlaneB,
    graph_client: GraphClient,
    store: TokenStore,
    approval: ApprovalEngine,
    member_email: str,
    account_id: str | None = None,
    clock: Callable[[], datetime] = _utcnow,
) -> dict:
    """US-3.4 — Obtém o chat 1:1 com uma pessoa, criando-o SE não existir (D1/D3).

    D3 — procurar primeiro: resolve o próprio email (`_own_email`); lê os chats existentes
    (paginando até `_MAX_LIST_FETCH`); procura um `oneOnOne` cujo único OUTRO membro == email
    (case-insensitive). Encontrado -> `{status:ok, chat_id, is_new_chat:false}` SEM token.
    Não encontrado (inclui membro sem email — A5) -> `pending_confirmation` (criar é escrita).
    A criação NÃO acontece aqui."""
    if not member_email:
        return {
            "status": "error",
            "message": "Indique o member_email (email já resolvido) do contacto.",
        }

    try:
        account, _ = await resolve_access_token(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            account_id=account_id, clock=clock,
        )
        own = await _own_email(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            graph_client=graph_client, account_id=account_id, clock=clock,
        )
        _, first = await call_graph(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            op=lambda token: graph_client.list_chats(token, top=50),
            account_id=account_id, clock=clock,
        )
        all_chats = list(first["chats"])
        next_link = first["next"]
        while next_link and len(all_chats) < _MAX_LIST_FETCH:
            async def fetch_more(token: str, _link: str = next_link) -> dict:
                return await graph_client.list_chats_next(token, _link)

            _, page = await call_graph(
                subject, mapping=mapping, plane_b=plane_b, store=store,
                op=fetch_more, account_id=account_id, clock=clock,
            )
            all_chats.extend(page["chats"])
            next_link = page["next"]
    except ReauthRequired as exc:
        return reauth_response(str(exc))

    target = member_email.casefold()
    own_cf = (own or "").casefold()
    existing = _find_one_on_one(all_chats, target=target, own=own_cf)
    if existing is not None:
        return {
            "status": "ok",
            "chat_id": existing,
            "is_new_chat": False,
            "message": "Já existe uma conversa 1:1 com este contacto.",
        }

    member_emails = [e for e in [own, member_email] if e]
    summary = f"Vai INICIAR uma nova conversa de Teams (1:1) com {member_email}."
    prepared = approval.prepare(
        subject=subject,
        account_id=account.account_id,
        operation="teams.chat_create",
        payload={"member_emails": member_emails, "member_email": member_email},
        summary=summary,
    )
    return prepared


def _find_one_on_one(chats: list[dict], *, target: str, own: str) -> str | None:
    """Devolve o id de um chat `oneOnOne` cujo único OUTRO membro (excluído o próprio) tem
    email == `target` (case-insensitive). None se não houver match (inclui o caso de o membro
    só trazer `userId` sem email — A5: segue-se para criação)."""
    for chat in chats:
        if chat.get("chat_type") != "oneOnOne":
            continue
        other_emails = [
            (m.get("email") or "").casefold()
            for m in (chat.get("members") or [])
            if (m.get("email") or "").casefold() != own
        ]
        if other_emails == [target]:
            return chat.get("id")
    return None


async def run_teams_get_or_create_one_on_one_chat_confirm(
    subject: str | None,
    *,
    mapping: IdentityMapping,
    plane_b: PlaneB,
    graph_client: GraphClient,
    store: TokenStore,
    approval: ApprovalEngine,
    confirmation_token: str,
    account_id: str | None = None,
    clock: Callable[[], datetime] = _utcnow,
) -> dict:
    """US-3.4 — Confirma a CRIAÇÃO da conversa 1:1 preparada (token fresco; idempotente no
    Graph); audita `teams.chat_create` com `extra={chat_type, is_new_chat}`."""
    async def executor(operation: str, payload: dict) -> dict:
        account, created = await call_graph(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            op=lambda token: graph_client.create_one_on_one_chat(
                token, member_emails=payload["member_emails"],
            ),
            account_id=account_id, clock=clock,
        )
        chat_id = (created or {}).get("id")
        log_audit(
            audit_logger, action="teams.chat_create", subject=subject,
            account_id=account.account_id, target=chat_id,
            outcome="success", recipients_count=1,
            extra={"chat_type": "oneOnOne", "is_new_chat": True},
        )
        return {
            "operation": operation,
            "chat_id": chat_id,
            "is_new_chat": True,
            "message": "Conversa iniciada.",
        }

    return await _confirm(approval, subject=subject, token=confirmation_token,
                          executor=executor)


async def _confirm(approval, *, subject, token, executor) -> dict:
    """Adaptador comum (idêntico ao email/calendário): traduz erros do ApprovalEngine e a
    reauth do executor em respostas amigáveis. Em ReauthRequired o token NÃO é consumido."""
    try:
        return await approval.confirm(subject=subject, token=token, executor=executor)
    except ReauthRequired as exc:
        return reauth_response(str(exc))
    except ConfirmationNotFound as exc:
        return {"status": "error", "message": str(exc)}
    except ConfirmationExpired as exc:
        return {"status": "expired", "message": str(exc)}
