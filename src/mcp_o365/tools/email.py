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
from datetime import datetime, timezone

from ..approval.engine import ApprovalEngine
from ..approval.errors import ConfirmationExpired, ConfirmationNotFound
from ..auth.errors import ReauthRequired
from ..auth.plane_b import PlaneB
from ..graph.client import GraphClient, recipients
from ..graph.sanitize import sanitize_html
from ..identity.mapping import IdentityMapping
from ..observability.audit import log_audit
from ..storage.token_store import TokenStore
from ._session import reauth_response, resolve_access_token

logger = logging.getLogger("mcp_o365.tools.email")
audit_logger = logging.getLogger("mcp_o365.audit")

# Limite de anexo inline (Graph): acima disto é preciso upload session.
_INLINE_ATTACHMENT_LIMIT = 3 * 1024 * 1024  # 3 MB

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
    top: int = 25,
    skip: int | None = None,
    account_id: str | None = None,
    clock: Callable[[], datetime] = _utcnow,
) -> dict:
    """US-1.1 — Pesquisa mensagens. Constrói `$search`/`$filter` a partir dos critérios."""
    try:
        _, access_token = await resolve_access_token(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            account_id=account_id, clock=clock,
        )
    except ReauthRequired as exc:
        return reauth_response(str(exc))

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

    result = await graph_client.list_messages(
        access_token,
        search=query,
        filter_query=filter_query,
        folder=folder,
        top=top,
        skip=skip,
        orderby=None if query else "receivedDateTime desc",
    )
    return {
        "status": "ok",
        "messages": result["messages"],
        "count": len(result["messages"]),
        "has_more": result["next"] is not None,
    }


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
        _, access_token = await resolve_access_token(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            account_id=account_id, clock=clock,
        )
    except ReauthRequired as exc:
        return reauth_response(str(exc))

    msg = await graph_client.get_message(access_token, message_id)
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
    account_id: str | None = None,
    clock: Callable[[], datetime] = _utcnow,
) -> dict:
    """US-1.5 — Lista anexos; com `download=True` + `attachment_id` devolve `contentBytes`."""
    try:
        _, access_token = await resolve_access_token(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            account_id=account_id, clock=clock,
        )
    except ReauthRequired as exc:
        return reauth_response(str(exc))

    if download and attachment_id:
        attachment = await graph_client.get_attachment(
            access_token, message_id, attachment_id
        )
        return {"status": "ok", "attachment": attachment}

    attachments = await graph_client.list_attachments(access_token, message_id)
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
    try:
        account, access_token = await resolve_access_token(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            account_id=account_id, clock=clock,
        )
    except ReauthRequired as exc:
        return reauth_response(str(exc))

    async def executor(operation: str, payload: dict) -> dict:
        message = payload["message"]
        if payload.get("large_attachments"):
            # Caminho de anexos grandes (>3MB): cria um rascunho com os anexos inline
            # (<=3MB), abre uma upload session por anexo grande, carrega os bytes em chunks
            # e só então envia o rascunho.
            inline = [
                a for a in message.get("attachments", [])
                if not _att_is_large(a)
            ]
            draft_msg = {**message, "attachments": inline}
            draft = await graph_client.create_draft(access_token, draft_msg)
            draft_id = draft.get("id")
            for att in message.get("attachments", []):
                if not _att_is_large(att):
                    continue
                raw = base64.b64decode(att.get("contentBytes") or "")
                session = await graph_client.create_attachment_upload_session(
                    access_token,
                    draft_id,
                    attachment_item={
                        "attachmentType": "file",
                        "name": att.get("name"),
                        "size": len(raw),
                    },
                )
                upload_url = session.get("uploadUrl") if session else None
                if upload_url:
                    await graph_client.upload_attachment_bytes(upload_url, raw)
            await graph_client.send_draft(access_token, draft_id)
        else:
            await graph_client.send_mail(access_token, message=message)
        log_audit(
            audit_logger,
            action="email.send",
            subject=subject,
            account_id=account.account_id,
            outcome="success",
            recipients_count=payload.get("recipients_count"),
            extra={"large_attachments": bool(payload.get("large_attachments"))},
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
    store: TokenStore,
    approval: ApprovalEngine,
    message_id: str,
    comment: str,
    mode: str = "reply",
    to_recipients: list[str] | None = None,
    account_id: str | None = None,
    clock: Callable[[], datetime] = _utcnow,
) -> dict:
    """US-1.4 — Prepara resposta/resposta-a-todos/reencaminho (mantém a thread)."""
    try:
        account, _ = await resolve_access_token(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            account_id=account_id, clock=clock,
        )
    except ReauthRequired as exc:
        return reauth_response(str(exc))

    if mode not in ("reply", "reply_all", "forward"):
        return {"status": "error", "message": f"Modo inválido: {mode}."}
    if mode == "forward" and not to_recipients:
        return {
            "status": "error",
            "message": "Reencaminhar exige destinatários (to_recipients).",
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
    try:
        account, access_token = await resolve_access_token(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            account_id=account_id, clock=clock,
        )
    except ReauthRequired as exc:
        return reauth_response(str(exc))

    async def executor(operation: str, payload: dict) -> dict:
        mode = payload["mode"]
        message_id = payload["message_id"]
        comment = payload["comment"]
        if mode == "forward":
            await graph_client.forward(
                access_token, message_id,
                comment=comment, to_recipients=payload["to_recipients"],
            )
            action = "email.forward"
            rcount = len(payload["to_recipients"])
        else:
            await graph_client.reply(
                access_token, message_id,
                comment=comment, reply_all=(mode == "reply_all"),
            )
            action = "email.reply"
            rcount = None
        log_audit(
            audit_logger, action=action, subject=subject,
            account_id=account.account_id, target=message_id,
            outcome="success", recipients_count=rcount, extra={"mode": mode},
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
    account_id: str | None = None,
    clock: Callable[[], datetime] = _utcnow,
) -> dict:
    """US-1.7 — Prepara mover: resolve nome de pasta -> id (case-insensitive)."""
    try:
        account, access_token = await resolve_access_token(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            account_id=account_id, clock=clock,
        )
    except ReauthRequired as exc:
        return reauth_response(str(exc))

    destination_id, display = await _resolve_destination_id(
        destination, access_token, graph_client
    )
    if destination_id is None:
        return {
            "status": "error",
            "message": f"Pasta de destino não encontrada: '{destination}'.",
        }

    return approval.prepare(
        subject=subject,
        account_id=account.account_id,
        operation="email.move",
        payload={"message_id": message_id, "destination_id": destination_id},
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
    try:
        account, access_token = await resolve_access_token(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            account_id=account_id, clock=clock,
        )
    except ReauthRequired as exc:
        return reauth_response(str(exc))

    async def executor(operation: str, payload: dict) -> dict:
        moved = await graph_client.move_message(
            access_token, payload["message_id"],
            destination_id=payload["destination_id"],
        )
        log_audit(
            audit_logger, action="email.move", subject=subject,
            account_id=account.account_id, target=payload["message_id"],
            outcome="success", extra={"destination_id": payload["destination_id"]},
        )
        return {
            "operation": operation,
            "message": "Mensagem movida.",
            "new_id": moved.get("id"),
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
        payload={"message_id": message_id, "permanent": permanent},
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
    try:
        account, access_token = await resolve_access_token(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            account_id=account_id, clock=clock,
        )
    except ReauthRequired as exc:
        return reauth_response(str(exc))

    # A recusa de eliminação permanente sem confirmação reforçada tem de acontecer ANTES
    # de consumir o token (idempotência) — espreita a operação pendente primeiro.
    pending = store.get_pending_operation(subject, confirmation_token)
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
        await graph_client.delete_message(access_token, payload["message_id"])
        log_audit(
            audit_logger, action="email.delete", subject=subject,
            account_id=account.account_id, target=payload["message_id"],
            outcome="success", extra={"permanent": bool(payload.get("permanent"))},
        )
        return {
            "operation": operation,
            "message": "Mensagem eliminada.",
            "permanent": bool(payload.get("permanent")),
        }

    return await _confirm(approval, subject=subject, token=confirmation_token,
                          executor=executor)


async def _confirm(approval, *, subject, token, executor) -> dict:
    """Adaptador comum: traduz erros do `ApprovalEngine` em respostas amigáveis."""
    try:
        return await approval.confirm(subject=subject, token=token, executor=executor)
    except ConfirmationNotFound as exc:
        return {"status": "error", "message": str(exc)}
    except ConfirmationExpired as exc:
        return {"status": "expired", "message": str(exc)}
