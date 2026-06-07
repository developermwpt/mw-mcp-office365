"""Fase 1 — Tools de Email (US-1.1 a US-1.8).

Funções `run_*` standalone, independentes do transporte MCP, para serem testáveis com Graph
e Entra mockados. Cada uma recebe o `subject` (do Plano A) e as dependências injetadas
(`mapping`, `plane_b`, `graph_client`, `store`, `approval`, `clock`).

Modelo de segurança:
- Resolução de token via `resolve_access_token`; qualquer `ReauthRequired` vira a resposta
  graciosa `reauth_required` (v1.1 §2.2), nunca exceção crua.
- Leituras (search/read/anexos) não exigem aprovação.
- Escritas (enviar/responder/reencaminhar/mover/eliminar) seguem o par prepare/confirm do
  `ApprovalEngine` (v1.1 §3): o `prepare` valida e monta; o `confirm` resolve um token FRESCO
  e só então executa, registando auditoria (só metadados — v1.1 §1.2).
- O corpo do email é conteúdo NÃO-confiável: ao ler, o HTML é sanitizado (`sanitize_html`)
  e marcado com `content_is_untrusted` para mitigar prompt injection (v1.1 §4).
"""

from __future__ import annotations

import base64
import logging
from collections.abc import Callable
from datetime import datetime, timedelta, timezone
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError  # só para apresentação (P3/§5)

from ..approval.engine import ApprovalEngine
from ..approval.errors import ConfirmationExpired, ConfirmationNotFound
from ..auth.errors import ReauthRequired
from ..auth.plane_b import PlaneB
from ..graph.attachments import extract_attachment_text
from ..graph.client import GraphClient, recipients
from ..graph.sanitize import sanitize_html
from ..identity.mapping import IdentityMapping
from ..observability.audit import log_audit
from ..storage.token_store import TokenStore
from ._session import call_graph, reauth_response, resolve_access_token
from ._timezone import _resolve_tz  # extraído de calendar.py (P6/§5.3)
from .learning import record_action_event

logger = logging.getLogger("mcp_o365.tools.email")
audit_logger = logging.getLogger("mcp_o365.audit")

# Limite de anexo inline (Graph): acima disto é preciso upload session.
_INLINE_ATTACHMENT_LIMIT = 3 * 1024 * 1024  # 3 MB

# Extended property MAPI PidTagDeferredSendTime (envio diferido nativo do Exchange).
# id = "{graph_type} {proptag}"; PT_SYSTIME -> SystemTime; tag 0x3FEF. Valor SEMPRE UTC ISO 8601.
_DEFERRED_SEND_PROP_ID = "SystemTime 0x3FEF"

# Validação temporal do agendamento (P1/P2).
_MIN_SCHEDULE_MARGIN = timedelta(minutes=2)    # >= 2 min no futuro
_MAX_SCHEDULE_HORIZON = timedelta(days=365)    # <= 1 ano

# Nomes de pastas bem-conhecidas (case-insensitive) -> id especial do Graph.
_WELL_KNOWN_FOLDERS = {
    "inbox": "inbox",
    "archive": "archive",
    "deleteditems": "deleteditems",
    "itens eliminados": "deleteditems",
    "arquivo": "archive",
    "caixa de entrada": "inbox",
    "drafts": "drafts",
    "sentitems": "sentitems",
    "junkemail": "junkemail",
}


def _utcnow() -> datetime:
    return datetime.now(timezone.utc)


# --- Pesquisa por período (US-1.1) ---
# Amplitude (em horas) até à qual se devolvem SEMPRE todos os emails do período sem
# perguntar. Acima disto, com mais resultados do que cabe numa página, pergunta-se primeiro.
_PERIOD_FETCH_ALL_HOURS = 24
# Teto de segurança ao paginar "todos" — evita puxar dezenas de milhar de emails.
_MAX_FETCH_ALL = 1000


def _parse_iso(value: str | None) -> datetime | None:
    """Parse de um instante ISO 8601 (aceita sufixo 'Z' e datas só-dia). None se inválido."""
    if not value:
        return None
    try:
        dt = datetime.fromisoformat(value.strip().replace("Z", "+00:00"))
    except ValueError:
        return None
    return dt if dt.tzinfo else dt.replace(tzinfo=timezone.utc)


def _period_within_a_day(
    date_from: str | None, date_to: str | None, clock: Callable[[], datetime]
) -> bool:
    """True se o período pedido estiver bem definido e for <= 24h.

    Exige limite inferior (`date_from`); o superior, se ausente, assume-se o instante atual.
    Datas inválidas ou ordem invertida (fim antes do início) -> False (trata como longo)."""
    start = _parse_iso(date_from)
    if start is None:
        return False
    end = _parse_iso(date_to) or clock()
    return timedelta(0) <= (end - start) <= timedelta(hours=_PERIOD_FETCH_ALL_HOURS)


# Metadados de mensagem permitidos no payload de aprovação (para a aprendizagem). NUNCA o
# corpo — coerente com a fronteira anti prompt injection (v1.1 §4) e o só-metadados.
_SAFE_META_KEYS = (
    "from", "sender", "subject", "hasAttachments", "importance", "isReply",
    "inReplyTo", "conversationId", "internetMessageHeaders", "is_newsletter", "list_id",
)


def _safe_meta(message_meta: dict | None) -> dict | None:
    """Filtra metadados de mensagem para o payload de aprovação (sem corpo nem PII extra)."""
    if not message_meta:
        return None
    return {k: message_meta[k] for k in _SAFE_META_KEYS if k in message_meta}


def _domains(addresses: list[str]) -> list[str]:
    """Domínios (sem PII de endereço) para auditoria — ordenados e únicos."""
    out = set()
    for a in addresses:
        if "@" in a:
            out.add(a.rsplit("@", 1)[1].lower())
    return sorted(out)


def _build_message(
    *,
    to: list[str],
    cc: list[str] | None,
    bcc: list[str] | None,
    subject: str,
    body: str,
    body_type: str,
    attachments: list[dict] | None,
) -> dict:
    """Monta o objeto `message` no formato Graph a partir dos campos validados."""
    message: dict = {
        "subject": subject,
        "body": {"contentType": body_type, "content": body},
        "toRecipients": recipients(to),
    }
    if cc:
        message["ccRecipients"] = recipients(cc)
    if bcc:
        message["bccRecipients"] = recipients(bcc)
    if attachments:
        message["attachments"] = [
            {
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": att["name"],
                "contentType": att.get("contentType", "application/octet-stream"),
                "contentBytes": att["contentBytes"],
                # `size` é preservado para que o confirm consiga distinguir anexos grandes
                # (>3MB) que têm de seguir o caminho de upload session, e não o inline.
                **({"size": att["size"]} if att.get("size") else {}),
            }
            for att in attachments
        ]
    return message


def _attachment_too_large(attachments: list[dict] | None) -> bool:
    """Indica se algum anexo excede o limite inline (precisa de upload session)."""
    for att in attachments or []:
        if att.get("size") and int(att["size"]) > _INLINE_ATTACHMENT_LIMIT:
            return True
    return False


# --- Agendamento (US-1.9/1.10/1.11): validação de offset e apresentação de hora ---


def _has_offset(value: str) -> bool:
    """True se a string ISO traz fuso explícito (Z ou ±HH:MM). P3: send_at TEM de o trazer."""
    v = value.strip()
    return v.endswith("Z") or v.endswith("z") or ("T" in v and ("+" in v[10:] or "-" in v[10:]))


def _to_utc_iso(dt: datetime) -> str:
    """datetime aware -> 'YYYY-MM-DDTHH:MM:SSZ' em UTC (formato aceite pela propriedade)."""
    return dt.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _present_in_tz(dt: datetime, tz_label: str | None) -> str:
    """Apresentação humana da hora no fuso (best-effort). NUNCA interpreta — só formata.

    Aceita IANA ('Europe/Lisbon') via zoneinfo; se for nome Windows ('GMT Standard Time')
    ou inválido, cai para UTC e anota. P3/§5: não é caminho crítico (o instante absoluto
    correto já está garantido pelo send_at com offset)."""
    if tz_label:
        try:
            local = dt.astimezone(ZoneInfo(tz_label))
            return f"{local.strftime('%d/%m/%Y %H:%M')} ({tz_label})"
        except (ZoneInfoNotFoundError, ValueError, KeyError):
            pass  # nome Windows ou inválido -> UTC
    return f"{dt.astimezone(timezone.utc).strftime('%d/%m/%Y %H:%M')} (UTC)"


def _extract_deferred_value(msg: dict) -> str | None:
    """Lê o value da extended property PidTagDeferredSendTime de um get_message expandido."""
    for ep in msg.get("singleValueExtendedProperties") or []:
        if ep.get("id") == _DEFERRED_SEND_PROP_ID:
            return ep.get("value")
    return None


# ============================ LEITURA (sem aprovação) ============================


async def run_email_search(
    subject: str | None,
    *,
    mapping: IdentityMapping,
    plane_b: PlaneB,
    graph_client: GraphClient,
    store: TokenStore,
    from_: str | None = None,
    subject_contains: str | None = None,
    date_from: str | None = None,
    date_to: str | None = None,
    query: str | None = None,
    folder: str | None = None,
    top: int = 50,
    skip: int | None = None,
    fetch_all: bool = False,
    account_id: str | None = None,
    clock: Callable[[], datetime] = _utcnow,
) -> dict:
    """US-1.1 — Pesquisa mensagens. Constrói `$search`/`$filter` a partir dos critérios.

    **Pesquisa por período** (date_from/date_to): quando há mais resultados do que cabe numa
    página (`top`), o comportamento depende da amplitude do período pedido:
    - **<= 24h**: devolve TODOS os emails do período (pagina automaticamente, com o teto de
      segurança `_MAX_FETCH_ALL`), sem perguntar.
    - **> 24h**: devolve `status='needs_clarification'` (sem paginar) para o assistente
      perguntar se quer todos (repetir com `fetch_all=True`) ou só os primeiros `top`.
    Sem período pedido, mantém-se o comportamento simples: primeira página + `has_more`."""
    filters: list[str] = []
    if from_:
        filters.append(f"from/emailAddress/address eq '{from_}'")
    if subject_contains:
        filters.append(f"contains(subject,'{subject_contains}')")
    if date_from:
        filters.append(f"receivedDateTime ge {date_from}")
    if date_to:
        filters.append(f"receivedDateTime le {date_to}")
    filter_query = " and ".join(filters) if filters else None
    orderby = None if query else "receivedDateTime desc"

    async def fetch_first(token: str) -> dict:
        return await graph_client.list_messages(
            token, search=query, filter_query=filter_query, folder=folder,
            top=top, skip=skip, orderby=orderby,
        )

    try:
        _, first = await call_graph(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            op=fetch_first, account_id=account_id, clock=clock,
        )
    except ReauthRequired as exc:
        return reauth_response(str(exc))

    has_more = first["next"] is not None
    has_period = bool(date_from or date_to)

    # Sem período pedido, ou tudo coube na 1ª página: comportamento simples de sempre.
    if not has_period or not has_more:
        return {
            "status": "ok",
            "messages": first["messages"],
            "count": len(first["messages"]),
            "has_more": has_more,
        }

    short_period = _period_within_a_day(date_from, date_to, clock)

    # Período > 24h com mais resultados: perguntar antes de puxar tudo (a menos que já
    # tenham pedido explicitamente fetch_all). NÃO assumir — devolve já a 1ª página.
    if not short_period and not fetch_all:
        return {
            "status": "needs_clarification",
            "question": (
                f"Encontrei mais de {top} emails neste período. Quer ver todos, "
                f"ou apenas os primeiros {top}?"
            ),
            "messages": first["messages"],
            "count": len(first["messages"]),
            "has_more": True,
            "options": [
                {
                    "label": f"Apenas os primeiros {top}",
                    "action": "use os resultados já devolvidos; não repita a pesquisa",
                },
                {
                    "label": "Todos",
                    "action": "repita email_search com fetch_all=true e os mesmos filtros",
                },
            ],
        }

    # Período <= 24h (devolve sempre todos) OU fetch_all pedido: pagina até esgotar (teto).
    all_messages = list(first["messages"])
    next_link = first["next"]
    truncated = False
    try:
        while next_link:
            if len(all_messages) >= _MAX_FETCH_ALL:
                truncated = True
                break

            async def fetch_more(token: str, _link: str = next_link) -> dict:
                return await graph_client.list_messages_next(
                    token, _link, consistency=bool(query)
                )

            _, page = await call_graph(
                subject, mapping=mapping, plane_b=plane_b, store=store,
                op=fetch_more, account_id=account_id, clock=clock,
            )
            all_messages.extend(page["messages"])
            next_link = page["next"]
    except ReauthRequired as exc:
        return reauth_response(str(exc))

    result = {
        "status": "ok",
        "messages": all_messages,
        "count": len(all_messages),
        "has_more": bool(next_link),
        "fetched_all": not truncated,
        "auto_fetched_all": short_period,
    }
    if truncated:
        result["truncated_at"] = _MAX_FETCH_ALL
    return result


async def run_email_read(
    subject: str | None,
    *,
    mapping: IdentityMapping,
    plane_b: PlaneB,
    graph_client: GraphClient,
    store: TokenStore,
    message_id: str,
    account_id: str | None = None,
    clock: Callable[[], datetime] = _utcnow,
) -> dict:
    """US-1.2 — Lê uma mensagem. Sanitiza o corpo HTML (conteúdo não-confiável)."""
    try:
        _, msg = await call_graph(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            op=lambda token: graph_client.get_message(token, message_id),
            account_id=account_id, clock=clock,
        )
    except ReauthRequired as exc:
        return reauth_response(str(exc))

    body = msg.get("body") or {}
    if (body.get("contentType") or "").lower() == "html" and body.get("content"):
        body = {**body, "content": sanitize_html(body["content"])}
    return {
        "status": "ok",
        "message": {**msg, "body": body},
        "content_is_untrusted": True,
    }


async def run_email_list_attachments(
    subject: str | None,
    *,
    mapping: IdentityMapping,
    plane_b: PlaneB,
    graph_client: GraphClient,
    store: TokenStore,
    message_id: str,
    download: bool = False,
    attachment_id: str | None = None,
    include_bytes: bool = False,
    account_id: str | None = None,
    clock: Callable[[], datetime] = _utcnow,
) -> dict:
    """US-1.5 — Lista anexos; com `download=True` + `attachment_id` devolve o **texto
    extraído no servidor** (PDF, Word .docx, PowerPoint .pptx e ficheiros de texto) pronto
    a ler. Para estes tipos NÃO use `include_bytes`: o conteúdo legível já vem em
    `extracted_text`. Os bytes em base64 só seguem com `include_bytes=True` (último recurso
    para tipos sem extração suportada)."""
    if download and attachment_id:
        try:
            _, attachment = await call_graph(
                subject, mapping=mapping, plane_b=plane_b, store=store,
                op=lambda token: graph_client.get_attachment(
                    token, message_id, attachment_id
                ),
                account_id=account_id, clock=clock,
            )
        except ReauthRequired as exc:
            return reauth_response(str(exc))
        meta = {
            k: attachment.get(k)
            for k in ("id", "name", "contentType", "size", "isInline")
        }
        extraction = extract_attachment_text(
            name=attachment.get("name"),
            content_type=attachment.get("contentType"),
            content_bytes_b64=attachment.get("contentBytes"),
        )
        result: dict = {
            "status": "ok",
            "attachment": meta,
            "content_is_untrusted": True,
        }
        if extraction.get("extractable"):
            result["extracted_text"] = extraction["text"]
            result["text_truncated"] = extraction.get("truncated", False)
            if "pages" in extraction:
                result["pages"] = extraction["pages"]
        else:
            result["extracted_text"] = None
            result["extraction_note"] = extraction.get("reason")
        # Os bytes só seguem se explicitamente pedidos (evita despejar base64 no contexto).
        if include_bytes:
            result["contentBytes"] = attachment.get("contentBytes")
        return result

    try:
        _, attachments = await call_graph(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            op=lambda token: graph_client.list_attachments(token, message_id),
            account_id=account_id, clock=clock,
        )
    except ReauthRequired as exc:
        return reauth_response(str(exc))
    return {"status": "ok", "attachments": attachments, "count": len(attachments)}


# ====================== ESCRITA (prepare / confirm) ======================


async def run_email_send_prepare(
    subject: str | None,
    *,
    mapping: IdentityMapping,
    plane_b: PlaneB,
    store: TokenStore,
    approval: ApprovalEngine,
    to: list[str],
    body: str,
    subject_line: str = "",
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    body_type: str = "Text",
    attachments: list[dict] | None = None,
    message_meta: dict | None = None,
    account_id: str | None = None,
    clock: Callable[[], datetime] = _utcnow,
) -> dict:
    """US-1.3 + US-1.6 — Prepara o envio: valida, monta a mensagem, devolve token."""
    try:
        account, _ = await resolve_access_token(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            account_id=account_id, clock=clock,
        )
    except ReauthRequired as exc:
        return reauth_response(str(exc))

    if not to:
        return {"status": "error", "message": "É obrigatório indicar destinatários (to)."}

    message = _build_message(
        to=to, cc=cc, bcc=bcc, subject=subject_line, body=body,
        body_type=body_type, attachments=attachments,
    )
    total = len(to) + len(cc or []) + len(bcc or [])
    large = _attachment_too_large(attachments)
    summary = (
        f"Enviar email para {total} destinatário(s) "
        f"(domínios: {', '.join(_domains(to)) or 'n/d'}), "
        f"assunto '{subject_line or '(sem assunto)'}'."
    )
    if large:
        summary += " Inclui anexo(s) grande(s) (envio via upload session)."

    prepared = approval.prepare(
        subject=subject,
        account_id=account.account_id,
        operation="email.send",
        payload={
            "message": message,
            "recipients_count": total,
            "large_attachments": large,
            "message_meta": _safe_meta(message_meta),
        },
        summary=summary,
    )
    prepared["recipients_count"] = total
    prepared["large_attachments"] = large
    return prepared


async def run_email_send_confirm(
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
    """US-1.3 + US-1.6 — Confirma o envio com token fresco; audita `email.send`."""
    async def executor(operation: str, payload: dict) -> dict:
        message = payload["message"]

        async def send(token: str) -> None:
            if payload.get("large_attachments"):
                # Anexos grandes (>3MB): rascunho com os inline (<=3MB), upload session por
                # anexo grande (bytes em chunks) e só então envia o rascunho.
                inline = [
                    a for a in message.get("attachments", [])
                    if not _att_is_large(a)
                ]
                draft = await graph_client.create_draft(
                    token, {**message, "attachments": inline}
                )
                draft_id = draft.get("id")
                for att in message.get("attachments", []):
                    if not _att_is_large(att):
                        continue
                    raw = base64.b64decode(att.get("contentBytes") or "")
                    session = await graph_client.create_attachment_upload_session(
                        token, draft_id,
                        attachment_item={
                            "attachmentType": "file",
                            "name": att.get("name"),
                            "size": len(raw),
                        },
                    )
                    upload_url = session.get("uploadUrl") if session else None
                    if upload_url:
                        await graph_client.upload_attachment_bytes(upload_url, raw)
                await graph_client.send_draft(token, draft_id)
            else:
                await graph_client.send_mail(token, message=message)

        account, _ = await call_graph(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            op=send, account_id=account_id, clock=clock,
        )
        log_audit(
            audit_logger,
            action="email.send",
            subject=subject,
            account_id=account.account_id,
            outcome="success",
            recipients_count=payload.get("recipients_count"),
            extra={"large_attachments": bool(payload.get("large_attachments"))},
        )
        record_action_event(
            subject, store=store, action="send",
            message=payload.get("message_meta"), clock=clock,
        )
        return {"operation": operation, "message": "Email enviado."}

    return await _confirm(approval, subject=subject, token=confirmation_token,
                          executor=executor)


def _att_is_large(att: dict) -> bool:
    return bool(att.get("size")) and int(att["size"]) > _INLINE_ATTACHMENT_LIMIT


async def run_email_reply_prepare(
    subject: str | None,
    *,
    mapping: IdentityMapping,
    plane_b: PlaneB,
    graph_client: GraphClient,
    store: TokenStore,
    approval: ApprovalEngine,
    message_id: str,
    comment: str,
    mode: str = "reply",
    to_recipients: list[str] | None = None,
    scope_confirmed: bool = False,
    message_meta: dict | None = None,
    account_id: str | None = None,
    clock: Callable[[], datetime] = _utcnow,
) -> dict:
    """US-1.4 — Prepara resposta/resposta-a-todos/reencaminho (mantém a thread).

    Quando `mode="reply"` e o email original tem vários destinatários (responder-a-todos
    alcançaria mais pessoas), devolve `needs_clarification` SEM criar token — para o assistente
    perguntar ao utilizador se quer responder só ao remetente ou a todos. `scope_confirmed=True`
    salta essa pergunta (o utilizador já escolheu "só ao remetente")."""
    if mode not in ("reply", "reply_all", "forward"):
        return {"status": "error", "message": f"Modo inválido: {mode}."}
    if mode == "forward" and not to_recipients:
        return {
            "status": "error",
            "message": "Reencaminhar exige destinatários (to_recipients).",
        }

    try:
        account, _ = await resolve_access_token(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            account_id=account_id, clock=clock,
        )
    except ReauthRequired as exc:
        return reauth_response(str(exc))

    # Desambiguação reply vs reply_all: se o email tem vários destinatários e o utilizador
    # apenas pediu "responder", pergunta o âmbito antes de preparar (não cria token).
    if mode == "reply" and not scope_confirmed:
        try:
            _, original = await call_graph(
                subject, mapping=mapping, plane_b=plane_b, store=store,
                op=lambda token: graph_client.get_message(token, message_id),
                account_id=account_id, clock=clock,
            )
        except ReauthRequired as exc:
            return reauth_response(str(exc))
        n_recipients = (
            len(original.get("toRecipients") or [])
            + len(original.get("ccRecipients") or [])
        )
        if n_recipients > 1:
            return {
                "status": "needs_clarification",
                "question": (
                    "Este email tem vários destinatários. Quer responder apenas a quem "
                    "enviou, ou a todos os destinatários?"
                ),
                "recipients_in_thread": n_recipients,
                "options": [
                    {
                        "label": "Apenas ao remetente",
                        "action": "repita email_reply_prepare com mode='reply' e "
                                  "scope_confirmed=true",
                    },
                    {
                        "label": "A todos",
                        "action": "repita email_reply_prepare com mode='reply_all'",
                    },
                ],
            }

    labels = {"reply": "Responder", "reply_all": "Responder a todos", "forward": "Reencaminhar"}
    summary = f"{labels[mode]} à mensagem {message_id}."
    if mode == "forward":
        summary += (
            f" Para {len(to_recipients or [])} destinatário(s) "
            f"(domínios: {', '.join(_domains(to_recipients or [])) or 'n/d'})."
        )

    return approval.prepare(
        subject=subject,
        account_id=account.account_id,
        operation="email.reply",
        payload={
            "message_id": message_id,
            "comment": comment,
            "mode": mode,
            "to_recipients": to_recipients or [],
            "message_meta": _safe_meta(message_meta),
        },
        summary=summary,
    )


async def run_email_reply_confirm(
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
    """US-1.4 — Confirma resposta/reencaminho; audita `email.reply`/`email.forward`."""
    async def executor(operation: str, payload: dict) -> dict:
        mode = payload["mode"]
        message_id = payload["message_id"]
        comment = payload["comment"]

        async def do(token: str) -> None:
            if mode == "forward":
                await graph_client.forward(
                    token, message_id,
                    comment=comment, to_recipients=payload["to_recipients"],
                )
            else:
                await graph_client.reply(
                    token, message_id,
                    comment=comment, reply_all=(mode == "reply_all"),
                )

        account, _ = await call_graph(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            op=do, account_id=account_id, clock=clock,
        )
        action = "email.forward" if mode == "forward" else "email.reply"
        rcount = len(payload["to_recipients"]) if mode == "forward" else None
        log_audit(
            audit_logger, action=action, subject=subject,
            account_id=account.account_id, target=message_id,
            outcome="success", recipients_count=rcount, extra={"mode": mode},
        )
        record_action_event(
            subject, store=store, action=mode,
            message=payload.get("message_meta"), clock=clock,
        )
        return {"operation": operation, "mode": mode, "message": "Operação concluída."}

    return await _confirm(approval, subject=subject, token=confirmation_token,
                          executor=executor)


async def run_email_move_prepare(
    subject: str | None,
    *,
    mapping: IdentityMapping,
    plane_b: PlaneB,
    graph_client: GraphClient,
    store: TokenStore,
    approval: ApprovalEngine,
    message_id: str,
    destination: str,
    message_meta: dict | None = None,
    account_id: str | None = None,
    clock: Callable[[], datetime] = _utcnow,
) -> dict:
    """US-1.7 — Prepara mover: resolve nome de pasta -> id (case-insensitive)."""
    try:
        account, (destination_id, display) = await call_graph(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            op=lambda token: _resolve_destination_id(destination, token, graph_client),
            account_id=account_id, clock=clock,
        )
    except ReauthRequired as exc:
        return reauth_response(str(exc))

    if destination_id is None:
        return {
            "status": "error",
            "message": f"Pasta de destino não encontrada: '{destination}'.",
        }

    return approval.prepare(
        subject=subject,
        account_id=account.account_id,
        operation="email.move",
        payload={
            "message_id": message_id,
            "destination_id": destination_id,
            "destination_name": display,
            "message_meta": _safe_meta(message_meta),
        },
        summary=f"Mover a mensagem {message_id} para a pasta '{display}'.",
    )


async def _resolve_destination_id(
    destination: str, access_token: str, graph_client: GraphClient
) -> tuple[str | None, str]:
    """Resolve nome de pasta -> (id, nome a apresentar). Aceita ids/bem-conhecidas."""
    key = destination.strip().lower()
    if key in _WELL_KNOWN_FOLDERS:
        return _WELL_KNOWN_FOLDERS[key], destination
    folders = await graph_client.list_folders(access_token)
    for f in folders:
        if (f.get("displayName") or "").lower() == key:
            return f["id"], f["displayName"]
        if f.get("id") == destination:
            return f["id"], f.get("displayName") or destination
    return None, destination


async def run_email_move_confirm(
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
    """US-1.7 — Confirma mover; audita `email.move`."""
    async def executor(operation: str, payload: dict) -> dict:
        account, moved = await call_graph(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            op=lambda token: graph_client.move_message(
                token, payload["message_id"],
                destination_id=payload["destination_id"],
            ),
            account_id=account_id, clock=clock,
        )
        log_audit(
            audit_logger, action="email.move", subject=subject,
            account_id=account.account_id, target=payload["message_id"],
            outcome="success", extra={"destination_id": payload["destination_id"]},
        )
        # Arquivar é um move para a pasta de arquivo — distinguir para a aprendizagem.
        behavior_action = (
            "archive" if payload["destination_id"] == "archive" else "move"
        )
        record_action_event(
            subject, store=store, action=behavior_action,
            message=payload.get("message_meta"),
            destination=payload.get("destination_name") or payload["destination_id"],
            clock=clock,
        )
        return {
            "operation": operation,
            "message": "Mensagem movida.",
            "new_id": (moved or {}).get("id"),
        }

    return await _confirm(approval, subject=subject, token=confirmation_token,
                          executor=executor)


async def run_email_delete_prepare(
    subject: str | None,
    *,
    mapping: IdentityMapping,
    plane_b: PlaneB,
    store: TokenStore,
    approval: ApprovalEngine,
    message_id: str,
    permanent: bool = False,
    message_meta: dict | None = None,
    account_id: str | None = None,
    clock: Callable[[], datetime] = _utcnow,
) -> dict:
    """US-1.8 — Prepara eliminação. Soft delete por defeito; permanente reforçada."""
    try:
        account, _ = await resolve_access_token(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            account_id=account_id, clock=clock,
        )
    except ReauthRequired as exc:
        return reauth_response(str(exc))

    if permanent:
        summary = (
            f"Eliminar PERMANENTEMENTE a mensagem {message_id} "
            "(irreversível). Requer confirmação reforçada."
        )
    else:
        summary = (
            f"Eliminar a mensagem {message_id} (vai para Itens Eliminados)."
        )

    prepared = approval.prepare(
        subject=subject,
        account_id=account.account_id,
        operation="email.delete",
        payload={
            "message_id": message_id,
            "permanent": permanent,
            "message_meta": _safe_meta(message_meta),
        },
        summary=summary,
    )
    if permanent:
        prepared["requires_reinforced_confirmation"] = True
    return prepared


async def run_email_delete_confirm(
    subject: str | None,
    *,
    mapping: IdentityMapping,
    plane_b: PlaneB,
    graph_client: GraphClient,
    store: TokenStore,
    approval: ApprovalEngine,
    confirmation_token: str,
    confirm_permanent: bool = False,
    account_id: str | None = None,
    clock: Callable[[], datetime] = _utcnow,
) -> dict:
    """US-1.8 — Confirma eliminação; permanente exige `confirm_permanent=True`."""
    # A recusa de eliminação permanente sem confirmação reforçada tem de acontecer ANTES
    # de consumir o token (idempotência) — espreita a operação pendente primeiro.
    pending = store.get_pending_operation(subject, confirmation_token) if subject else None
    if (
        pending is not None
        and pending["consumed_at"] is None
        and pending["payload"].get("permanent")
        and not confirm_permanent
    ):
        return {
            "status": "error",
            "message": (
                "Eliminação permanente requer confirmação reforçada: "
                "repita com confirm_permanent=True."
            ),
        }

    async def executor(operation: str, payload: dict) -> dict:
        is_permanent = bool(payload.get("permanent"))

        async def do(token: str):
            if is_permanent:
                # Hard delete real: vai para a pasta `purges`, irrecuperável.
                return await graph_client.permanent_delete(token, payload["message_id"])
            # Soft delete previsível: mover explicitamente para Itens Eliminados (visível e
            # recuperável). Devolve o novo recurso (o id muda com o move).
            return await graph_client.move_message(
                token, payload["message_id"], destination_id="deleteditems"
            )

        account, moved = await call_graph(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            op=do, account_id=account_id, clock=clock,
        )
        log_audit(
            audit_logger, action="email.delete", subject=subject,
            account_id=account.account_id, target=payload["message_id"],
            outcome="success", extra={"permanent": is_permanent},
        )
        record_action_event(
            subject, store=store, action="delete",
            message=payload.get("message_meta"), clock=clock,
        )
        result = {
            "operation": operation,
            "message": (
                "Mensagem eliminada permanentemente."
                if is_permanent
                else "Mensagem movida para Itens Eliminados."
            ),
            "permanent": is_permanent,
        }
        if not is_permanent and isinstance(moved, dict) and moved.get("id"):
            result["new_id"] = moved["id"]
        return result

    return await _confirm(approval, subject=subject, token=confirmation_token,
                          executor=executor)


# ==================== AGENDAMENTO DE ENVIO (US-1.9/1.10/1.11) ====================


async def run_email_schedule_prepare(
    subject: str | None,
    *,
    mapping: IdentityMapping,
    plane_b: PlaneB,
    graph_client: GraphClient,
    store: TokenStore,
    approval: ApprovalEngine,
    to: list[str],
    body: str,
    send_at: str,
    subject_line: str = "",
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    body_type: str = "Text",
    attachments: list[dict] | None = None,
    timezone: str | None = None,
    message_meta: dict | None = None,
    account_id: str | None = None,
    clock: Callable[[], datetime] = _utcnow,
) -> dict:
    """US-1.9 — Prepara o agendamento: valida destinatários e a HORA, monta a mensagem com a
    extended property de envio diferido, devolve token. NÃO escreve no Graph."""
    try:
        account, _ = await resolve_access_token(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            account_id=account_id, clock=clock,
        )
    except ReauthRequired as exc:
        return reauth_response(str(exc))

    if not to:
        return {"status": "error", "message": "É obrigatório indicar destinatários (to)."}

    # --- Validação temporal (P1/P2) — ANTES de qualquer escrita, sem token. ---
    when = _parse_iso(send_at)                       # reutiliza o helper existente (aceita 'Z')
    if when is None:
        return {"status": "error", "message":
                "send_at inválido: indique um instante ISO 8601 com fuso/offset (ex.: "
                "2026-06-10T09:00:00+01:00 ou ...Z)."}
    # P3: exigir instante ABSOLUTO. _parse_iso completa naive->UTC, por isso validamos
    # explicitamente a presença de offset no texto.
    if not _has_offset(send_at):
        return {"status": "error", "message":
                "send_at tem de incluir o fuso/offset resolvido (ex.: +01:00 ou Z). "
                "Resolva a hora no fuso do mailbox a montante."}
    now = clock()
    delta = when - now
    if delta < _MIN_SCHEDULE_MARGIN:
        return {"status": "error", "message":
                "A hora de envio tem de estar pelo menos 2 minutos no futuro."}
    if delta > _MAX_SCHEDULE_HORIZON:
        return {"status": "error", "message":
                "A hora de envio não pode estar a mais de 1 ano de distância."}

    send_at_utc = _to_utc_iso(when)                  # normaliza para "...Z"

    # --- Fuso só para APRESENTAÇÃO (best-effort; P3/§5). Nunca interpreta a hora. ---
    tz_label = timezone or await _resolve_tz(
        subject, mapping=mapping, plane_b=plane_b, store=store,
        graph_client=graph_client, account_id=account_id, clock=clock,
    )  # pode ser None (fuso indisponível) -> apresentar em UTC e declará-lo.
    when_label = _present_in_tz(when, tz_label)

    # --- Montagem da message + extended property (mesmo _build_message do send). ---
    message = _build_message(
        to=to, cc=cc, bcc=bcc, subject=subject_line, body=body,
        body_type=body_type, attachments=attachments,
    )
    message["singleValueExtendedProperties"] = [
        {"id": _DEFERRED_SEND_PROP_ID, "value": send_at_utc}
    ]

    total = len(to) + len(cc or []) + len(bcc or [])
    large = _attachment_too_large(attachments)
    summary = (
        f"Agendar email para {total} destinatário(s) "
        f"(domínios: {', '.join(_domains(to)) or 'n/d'}), "
        f"assunto '{subject_line or '(sem assunto)'}', envio em {when_label}."
    )
    if large:
        summary += " Inclui anexo(s) grande(s) (envio via upload session)."
    if tz_label is None and timezone is None:
        summary += " (Fuso do mailbox indisponível; hora interpretada/apresentada em UTC.)"

    prepared = approval.prepare(
        subject=subject,
        account_id=account.account_id,
        operation="email.schedule",
        payload={
            "message": message,                      # já inclui singleValueExtendedProperties
            "recipients_count": total,
            "large_attachments": large,
            "send_at_utc": send_at_utc,              # para auditoria (só-metadados)
            "message_meta": _safe_meta(message_meta),
        },
        summary=summary,
    )
    prepared["recipients_count"] = total
    prepared["large_attachments"] = large
    prepared["send_at_utc"] = send_at_utc
    return prepared


async def run_email_schedule_confirm(
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
    """US-1.9 — Confirma o agendamento com token fresco; audita `email.schedule`."""
    async def executor(operation: str, payload: dict) -> dict:
        message = payload["message"]                 # inclui singleValueExtendedProperties

        async def do(token: str) -> str | None:
            # Diferido OBRIGA draft->send (a extended property é definida na CRIAÇÃO do
            # rascunho; `sendMail` ignora-a). Mesmo sem anexos grandes, usa-se draft->send.
            inline = [a for a in message.get("attachments", []) if not _att_is_large(a)]
            draft = await graph_client.create_draft(token, {**message, "attachments": inline})
            draft_id = draft.get("id")
            if payload.get("large_attachments"):
                for att in message.get("attachments", []):
                    if not _att_is_large(att):
                        continue
                    raw = base64.b64decode(att.get("contentBytes") or "")
                    session = await graph_client.create_attachment_upload_session(
                        token, draft_id,
                        attachment_item={
                            "attachmentType": "file",
                            "name": att.get("name"),
                            "size": len(raw),
                        },
                    )
                    upload_url = session.get("uploadUrl") if session else None
                    if upload_url:
                        await graph_client.upload_attachment_bytes(upload_url, raw)
            await graph_client.send_draft(token, draft_id)
            return draft_id

        account, draft_id = await call_graph(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            op=do, account_id=account_id, clock=clock,
        )
        log_audit(
            audit_logger, action="email.schedule", subject=subject,
            account_id=account.account_id, target=draft_id,
            outcome="success", recipients_count=payload.get("recipients_count"),
            extra={"large_attachments": bool(payload.get("large_attachments")),
                   "send_at_utc": payload.get("send_at_utc"), "deferred": True},
        )
        record_action_event(
            subject, store=store, action="schedule",      # P5
            message=payload.get("message_meta"), clock=clock,
        )
        return {"operation": operation, "message_id": draft_id,
                "send_at_utc": payload.get("send_at_utc"),
                "message": "Envio agendado."}

    return await _confirm(approval, subject=subject, token=confirmation_token,
                          executor=executor)


async def run_email_list_scheduled(
    subject: str | None,
    *,
    mapping: IdentityMapping,
    plane_b: PlaneB,
    graph_client: GraphClient,
    store: TokenStore,
    top: int = 50,
    timezone: str | None = None,
    account_id: str | None = None,
    clock: Callable[[], datetime] = _utcnow,
) -> dict:
    """US-1.10 — Lista os rascunhos com a extended property de envio diferido cujo instante
    ainda é FUTURO (pendentes). Leitura; não escreve; não exige aprovação."""
    tz_label = timezone or await _resolve_tz(
        subject, mapping=mapping, plane_b=plane_b, store=store,
        graph_client=graph_client, account_id=account_id, clock=clock,
    )
    try:
        _, page = await call_graph(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            op=lambda token: graph_client.list_deferred_drafts(
                token, prop_id=_DEFERRED_SEND_PROP_ID, top=top,
            ),
            account_id=account_id, clock=clock,
        )
    except ReauthRequired as exc:
        return reauth_response(str(exc))

    now = clock()
    items = []
    for d in page["drafts"]:
        send_at_utc = d.get("deferred_send_at")          # valor da prop expandida (UTC ISO)
        when = _parse_iso(send_at_utc)
        if when is None or when <= now:                  # FILTRO "ainda futuro" CLIENT-SIDE
            continue
        recips = d.get("to") or []
        items.append({
            "id": d.get("id"),
            "subject": sanitize_html(d.get("subject") or ""),  # conteúdo não-confiável
            "recipients_count": len(recips),
            "recipient_domains": _domains(recips),
            "send_at": _present_in_tz(when, tz_label),   # apresentação no fuso
            "send_at_utc": send_at_utc,
        })
    items.sort(key=lambda i: i["send_at_utc"])            # mais próximos primeiro
    return {
        "status": "ok",
        "scheduled": items,
        "count": len(items),
        "has_more": page["next"] is not None,
        "content_is_untrusted": True,
    }


async def run_email_schedule_cancel_prepare(
    subject: str | None,
    *,
    mapping: IdentityMapping,
    plane_b: PlaneB,
    graph_client: GraphClient,
    store: TokenStore,
    approval: ApprovalEngine,
    message_id: str,
    timezone: str | None = None,
    message_meta: dict | None = None,
    account_id: str | None = None,
    clock: Callable[[], datetime] = _utcnow,
) -> dict:
    """US-1.11 — Prepara o cancelamento de um envio agendado (NÃO elimina). Confirma
    best-effort que o `message_id` é um rascunho ainda diferido e monta o resumo."""
    try:
        account, _ = await resolve_access_token(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            account_id=account_id, clock=clock,
        )
    except ReauthRequired as exc:
        return reauth_response(str(exc))

    # Best-effort (P7): ler o rascunho com a extended property para validar/apresentar.
    deferred_at = None
    subj = None
    try:
        _, msg = await call_graph(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            op=lambda token: graph_client.get_message(
                token, message_id,
                expand=(
                    "singleValueExtendedProperties("
                    f"$filter=id eq '{_DEFERRED_SEND_PROP_ID}')"
                ),
            ),
            account_id=account_id, clock=clock,
        )
        deferred_at = _extract_deferred_value(msg)       # None se a prop não estiver presente
        subj = msg.get("subject")
    except ReauthRequired as exc:
        return reauth_response(str(exc))
    # (Outras falhas de leitura: degradar — não bloquear o cancelamento; resumo genérico.)

    # Se lemos e o rascunho JÁ NÃO é diferido (sem a prop / já enviado) -> erro, sem token.
    if subj is not None and deferred_at is None:
        return {"status": "error", "message":
                "Este email já não é um envio agendado pendente (pode já ter sido enviado "
                "ou cancelado). Liste os agendados com email_list_scheduled."}

    tz_label = timezone or await _resolve_tz(
        subject, mapping=mapping, plane_b=plane_b, store=store,
        graph_client=graph_client, account_id=account_id, clock=clock,
    )
    when = _parse_iso(deferred_at)
    when_label = _present_in_tz(when, tz_label) if when else "hora desconhecida"
    summary = (
        f"Cancelar o envio agendado para {when_label}, "
        f"assunto '{sanitize_html(subj or '(desconhecido)')}'. "
        "O rascunho vai para Itens Eliminados (recuperável)."
    )

    return approval.prepare(
        subject=subject,
        account_id=account.account_id,
        operation="email.schedule_cancel",
        payload={"message_id": message_id, "message_meta": _safe_meta(message_meta)},
        summary=summary,
    )


async def run_email_schedule_cancel_confirm(
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
    """US-1.11 — Confirma o cancelamento; soft delete do rascunho; audita
    `email.schedule_cancel`."""
    async def executor(operation: str, payload: dict) -> dict:
        account, moved = await call_graph(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            op=lambda token: graph_client.move_message(
                token, payload["message_id"], destination_id="deleteditems"),
            account_id=account_id, clock=clock,
        )
        log_audit(
            audit_logger, action="email.schedule_cancel", subject=subject,
            account_id=account.account_id, target=payload["message_id"],
            outcome="success", extra={"permanent": False},
        )
        result = {"operation": operation, "message": "Envio agendado cancelado "
                  "(rascunho movido para Itens Eliminados)."}
        if isinstance(moved, dict) and moved.get("id"):
            result["new_id"] = moved["id"]
        return result

    return await _confirm(approval, subject=subject, token=confirmation_token,
                          executor=executor)


async def _confirm(approval, *, subject, token, executor) -> dict:
    """Adaptador comum: traduz erros do `ApprovalEngine` (e a reauth do executor) em
    respostas amigáveis. Em `ReauthRequired`, a operação NÃO é marcada como consumida pelo
    engine (o executor levantou antes de concluir), pelo que o token continua válido para
    repetir após o re-login."""
    try:
        return await approval.confirm(subject=subject, token=token, executor=executor)
    except ReauthRequired as exc:
        return reauth_response(str(exc))
    except ConfirmationNotFound as exc:
        return {"status": "error", "message": str(exc)}
    except ConfirmationExpired as exc:
        return {"status": "expired", "message": str(exc)}
