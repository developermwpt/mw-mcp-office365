"""Fase 2 — Tools de Calendário (US-2.1 a US-2.6).

Funções `run_calendar_*` standalone, independentes do transporte MCP, testáveis com Graph e
Entra mockados. Seguem à risca os padrões da Fase 1 (`tools/email.py`): resolução de token
via `resolve_access_token`/`call_graph`, reautenticação graciosa (`reauth_response`), o par
prepare/confirm do `ApprovalEngine` para escritas, auditoria só-metadados (`log_audit`) e
sanitização de conteúdo não-confiável (`sanitize_html`).

Decisões fechadas implementadas (D1–D9, ver docs/fase-2/plano-implementacao.md):
- D1 — fuso do MAILBOX, lido uma vez por pedido (`_resolve_tz`) e usado no header `Prefer`
  das leituras e nos `start/end.timeZone` das escritas.
- D2 — disponibilidade via `getSchedule` (próprio + participantes).
- D3 — escritas com participantes declaram no resumo "notifica N participantes (domínios: …)".
- D4 — só calendário primário; ler ocorrências expandidas; editar/cancelar série recorrente
  devolve `needs_clarification` (esta ocorrência vs série). NÃO cria séries.
- D5 — a consulta de eventos AUTO-PAGINA todo o intervalo (teto `_MAX_FETCH_ALL`), sem perguntar.
- D6 — sem `location` -> link Teams; com `location` -> presencial sem link; o resumo declara.
- D7 — responder lê o estado atual e declara a mudança; bloqueia se o subject for organizador.
- D9 — participantes chegam como emails já resolvidos (resolve_recipient é feito a montante).
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
from ..logging_setup import subject_hash
from ..observability.audit import log_audit
from ..storage.token_store import TokenStore
from ._session import call_graph, reauth_response, resolve_access_token
from .email import _domains

logger = logging.getLogger("mcp_o365.tools.calendar")
audit_logger = logging.getLogger("mcp_o365.audit")

# Teto de segurança da auto-paginação (D5), igual ao email.
_MAX_FETCH_ALL = 1000
# Respostas a convite aceites (chave amigável -> action do Graph).
_VALID_RESPONSES = {
    "accept": "accept",
    "decline": "decline",
    "tentative": "tentativelyAccept",
}

# Estado de resposta do Graph -> rótulo PT (para D7 declarar a transição).
_RESPONSE_PT = {
    "none": "Sem resposta",
    "notResponded": "Sem resposta",
    "organizer": "Organizador",
    "tentativelyAccepted": "Tentativo",
    "accepted": "Aceitado",
    "declined": "Recusado",
}
# Action do Graph -> rótulo PT da NOVA resposta.
_NEW_RESPONSE_PT = {
    "accept": "Aceitado",
    "decline": "Recusado",
    "tentativelyAccept": "Tentativo",
}


def _utcnow() -> datetime:
    return datetime.now(timezone.utc)


async def _resolve_tz(
    subject: str | None,
    *,
    mapping: IdentityMapping,
    plane_b: PlaneB,
    store: TokenStore,
    graph_client: GraphClient,
    account_id: str | None,
    clock: Callable[[], datetime],
) -> str | None:
    """D1 — lê o fuso do mailbox uma vez por pedido. None -> Graph usa UTC (sem `Prefer`).

    BEST-EFFORT por desenho: ler o fuso é acessório e exige o scope `MailboxSettings.Read`.
    Se esse scope faltar (Graph devolve 403 -> `UpstreamAuthError`) ou a sessão não resolver,
    devolve-se `None` (fallback para UTC) em vez de propagar. Nunca se passa por `call_graph`
    aqui, para que uma falha do fuso NÃO force refresh nem marque a conta como expirada —
    senão um 403 numa leitura secundária derrubava a sessão inteira (email incluído)."""
    try:
        _, token = await resolve_access_token(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            account_id=account_id, clock=clock,
        )
        return await graph_client.get_mailbox_timezone(token)
    except (UpstreamAuthError, ReauthRequired):
        # Fuso indisponível (scope em falta ou sessão por reautenticar) -> UTC.
        # A reautenticação genuína, se necessária, será sinalizada pela chamada principal
        # (ex.: calendarView), que usa scopes efetivamente concedidos.
        return None


async def _own_email(
    subject: str | None,
    *,
    mapping: IdentityMapping,
    plane_b: PlaneB,
    store: TokenStore,
    graph_client: GraphClient,
    account_id: str | None,
    clock: Callable[[], datetime],
) -> str | None:
    """Email do próprio utilizador (via `/me`) — usado em getSchedule e bloqueio de organizador."""
    _, me = await call_graph(
        subject, mapping=mapping, plane_b=plane_b, store=store,
        op=lambda token: graph_client.me(token),
        account_id=account_id, clock=clock,
    )
    return me.get("userPrincipalName")


def _sanitize_event_summary(event: dict) -> dict:
    """Sanitiza o `bodyPreview` de um evento resumido (conteúdo não-confiável), como no email.

    O `bodyPreview` é texto vindo do mailbox e pode conter tentativas de prompt injection;
    passa por `sanitize_html` antes de chegar ao modelo."""
    preview = event.get("bodyPreview")
    if preview:
        return {**event, "bodyPreview": sanitize_html(preview)}
    return event


# ============================ LEITURA (sem aprovação) ============================


async def run_calendar_list_events(
    subject: str | None,
    *,
    mapping: IdentityMapping,
    plane_b: PlaneB,
    graph_client: GraphClient,
    store: TokenStore,
    start: str,
    end: str,
    top: int = 50,
    account_id: str | None = None,
    clock: Callable[[], datetime] = _utcnow,
) -> dict:
    """US-2.1 — Lista eventos do calendário primário num intervalo (read-only).

    Resolve o fuso do mailbox (D1), lê a 1ª página de `calendarView` e AUTO-PAGINA todo o
    intervalo seguindo `@odata.nextLink` (D5), com o teto `_MAX_FETCH_ALL`. Ocorrências de
    séries vêm expandidas; o corpo/preview é conteúdo não-confiável (sanitizado/flag)."""
    if not start or not end:
        return {
            "status": "error",
            "message": "Indique o intervalo: start e end (ISO 8601).",
        }

    try:
        tz = await _resolve_tz(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            graph_client=graph_client, account_id=account_id, clock=clock,
        )

        async def fetch_first(token: str) -> dict:
            return await graph_client.list_calendar_view(
                token, start=start, end=end, top=top, prefer_timezone=tz,
            )

        _, first = await call_graph(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            op=fetch_first, account_id=account_id, clock=clock,
        )

        all_events = list(first["events"])
        next_link = first["next"]
        truncated = False
        while next_link:
            if len(all_events) >= _MAX_FETCH_ALL:
                truncated = True
                break

            async def fetch_more(token: str, _link: str = next_link) -> dict:
                return await graph_client.list_calendar_view_next(
                    token, _link, prefer_timezone=tz,
                )

            _, page = await call_graph(
                subject, mapping=mapping, plane_b=plane_b, store=store,
                op=fetch_more, account_id=account_id, clock=clock,
            )
            all_events.extend(page["events"])
            next_link = page["next"]
    except ReauthRequired as exc:
        return reauth_response(str(exc))

    all_events = [_sanitize_event_summary(e) for e in all_events]
    result = {
        "status": "ok",
        "events": all_events,
        "count": len(all_events),
        "timezone": tz,
        "has_more": bool(next_link),
        "fetched_all": not truncated,
        "auto_fetched_all": True,
        "content_is_untrusted": True,
    }
    if truncated:
        result["truncated_at"] = _MAX_FETCH_ALL
        result["fetched_all"] = False
        result["auto_fetched_all"] = False
    return result


async def run_calendar_check_availability(
    subject: str | None,
    *,
    mapping: IdentityMapping,
    plane_b: PlaneB,
    graph_client: GraphClient,
    store: TokenStore,
    attendees: list[str] | None = None,
    start: str,
    end: str,
    interval_minutes: int = 30,
    account_id: str | None = None,
    clock: Callable[[], datetime] = _utcnow,
) -> dict:
    """US-2.2 — Disponibilidade (livre/ocupado) do próprio + participantes (read-only, D2).

    `getSchedule` chamado uma vez; o próprio é SEMPRE incluído mesmo que não venha em
    `attendees`. Horas no fuso do mailbox."""
    if not start or not end:
        return {
            "status": "error",
            "message": "Indique o intervalo: start e end (ISO 8601).",
        }

    try:
        tz = await _resolve_tz(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            graph_client=graph_client, account_id=account_id, clock=clock,
        )
        own = await _own_email(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            graph_client=graph_client, account_id=account_id, clock=clock,
        )
        schedules: list[str] = []
        if own:
            schedules.append(own)
        for a in attendees or []:
            if a and a.lower() not in {s.lower() for s in schedules}:
                schedules.append(a)

        _, value = await call_graph(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            op=lambda token: graph_client.get_schedule(
                token, schedules=schedules, start=start, end=end,
                interval_minutes=interval_minutes, prefer_timezone=tz,
            ),
            account_id=account_id, clock=clock,
        )
    except ReauthRequired as exc:
        return reauth_response(str(exc))

    return {
        "status": "ok",
        "timezone": tz,
        "schedules": value,
        "interval_minutes": interval_minutes,
    }


# ====================== ESCRITA (prepare / confirm) ======================


async def run_calendar_create_prepare(
    subject: str | None,
    *,
    mapping: IdentityMapping,
    plane_b: PlaneB,
    graph_client: GraphClient,
    store: TokenStore,
    approval: ApprovalEngine,
    subject_line: str,
    start: str,
    end: str,
    attendees: list[str] | None = None,
    body: str = "",
    body_type: str = "Text",
    location: str | None = None,
    online: bool | None = None,
    account_id: str | None = None,
    clock: Callable[[], datetime] = _utcnow,
) -> dict:
    """US-2.3 — Prepara a criação de um evento (NÃO cria; só lê o fuso para montar start/end).

    D6: sem `location` -> link Teams; com `location` -> presencial sem link. D3: o resumo
    declara participantes notificados (domínios) e se inclui ou não link Teams."""
    if not subject_line or not start or not end:
        return {
            "status": "error",
            "message": "Indique subject_line, start e end (ISO 8601).",
        }

    try:
        account, _ = await resolve_access_token(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            account_id=account_id, clock=clock,
        )
        # Leitura apenas (resolver fuso) — o invariante "prepare não escreve" mantém-se.
        tz = await _resolve_tz(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            graph_client=graph_client, account_id=account_id, clock=clock,
        )
    except ReauthRequired as exc:
        return reauth_response(str(exc))

    online_meeting = online if online is not None else (location is None)
    recipients = attendees or []
    event = {
        "subject": subject_line,
        "start": {"dateTime": start, "timeZone": tz or "UTC"},
        "end": {"dateTime": end, "timeZone": tz or "UTC"},
        "body": {"contentType": body_type, "content": body},
        "attendees": [
            {"emailAddress": {"address": a}, "type": "required"} for a in recipients
        ],
        **({"location": {"displayName": location}} if location else {}),
        **(
            {"isOnlineMeeting": True, "onlineMeetingProvider": "teamsForBusiness"}
            if online_meeting
            else {}
        ),
    }

    n = len(recipients)
    teams_phrase = (
        "Inclui link Teams."
        if online_meeting
        else f"Presencial em '{location}' (sem link Teams)."
        if location
        else "Sem link Teams."
    )
    summary = (
        f"Criar evento '{subject_line}' de {start} a {end} ({tz or 'UTC'}). "
        f"Notifica {n} participante(s) (domínios: {', '.join(_domains(recipients)) or 'n/d'}). "
        f"{teams_phrase}"
    )

    prepared = approval.prepare(
        subject=subject,
        account_id=account.account_id,
        operation="calendar.create",
        payload={
            "event": event,
            "recipients_count": n,
            "online": online_meeting,
            "event_subject": subject_line,
        },
        summary=summary,
    )
    prepared["recipients_count"] = n
    return prepared


async def run_calendar_create_confirm(
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
    """US-2.3 — Confirma e cria o evento preparado (token fresco); audita `calendar.create`."""
    async def executor(operation: str, payload: dict) -> dict:
        account, created = await call_graph(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            op=lambda token: graph_client.create_event(token, event=payload["event"]),
            account_id=account_id, clock=clock,
        )
        log_audit(
            audit_logger, action="calendar.create", subject=subject,
            account_id=account.account_id, target=(created or {}).get("id"),
            outcome="success", recipients_count=payload.get("recipients_count"),
            extra={
                "subject_hash": subject_hash(payload.get("event_subject") or ""),
                "online": bool(payload.get("online")),
            },
        )
        return {
            "operation": operation,
            "event_id": (created or {}).get("id"),
            "web_link": (created or {}).get("webLink"),
            "message": "Evento criado.",
        }

    return await _confirm(approval, subject=subject, token=confirmation_token,
                          executor=executor)


def _recurrence_clarification(question_verb: str) -> dict:
    """D4 — resposta `needs_clarification` para editar/cancelar série recorrente."""
    tool = "calendar_update_prepare" if question_verb == "Editar" else "calendar_cancel_prepare"
    return {
        "status": "needs_clarification",
        "question": (
            f"Este evento é recorrente. {question_verb} só esta ocorrência, "
            "ou a série inteira?"
        ),
        "options": [
            {
                "label": "Só esta ocorrência",
                "action": f"repita {tool} com scope='occurrence'",
            },
            {
                "label": "A série inteira",
                "action": f"repita {tool} com scope='series'",
            },
        ],
    }


async def run_calendar_update_prepare(
    subject: str | None,
    *,
    mapping: IdentityMapping,
    plane_b: PlaneB,
    graph_client: GraphClient,
    store: TokenStore,
    approval: ApprovalEngine,
    event_id: str,
    subject_line: str | None = None,
    start: str | None = None,
    end: str | None = None,
    location: str | None = None,
    body: str | None = None,
    body_type: str = "Text",
    attendees: list[str] | None = None,
    scope: str | None = None,
    account_id: str | None = None,
    clock: Callable[[], datetime] = _utcnow,
) -> dict:
    """US-2.4 — Prepara editar/reagendar (NÃO altera). Recorrente sem `scope` -> clarification.

    `scope='occurrence'` -> PATCH ao próprio `event_id`; `scope='series'` -> PATCH ao
    `seriesMasterId`. Só os campos não-None entram em `changes`."""
    if not event_id:
        return {"status": "error", "message": "Indique o event_id a editar."}

    try:
        account, _ = await resolve_access_token(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            account_id=account_id, clock=clock,
        )
        tz = await _resolve_tz(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            graph_client=graph_client, account_id=account_id, clock=clock,
        )
        _, event = await call_graph(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            op=lambda token: graph_client.get_event(token, event_id, prefer_timezone=tz),
            account_id=account_id, clock=clock,
        )
    except ReauthRequired as exc:
        return reauth_response(str(exc))

    if event.get("isRecurring") and scope is None:
        return _recurrence_clarification("Editar")

    target_event_id = event_id
    if scope == "series" and event.get("seriesMasterId"):
        target_event_id = event["seriesMasterId"]

    changes: dict = {}
    changed_fields: list[str] = []
    if subject_line is not None:
        changes["subject"] = subject_line
        changed_fields.append("assunto")
    if start is not None:
        changes["start"] = {"dateTime": start, "timeZone": tz or "UTC"}
        changed_fields.append("início")
    if end is not None:
        changes["end"] = {"dateTime": end, "timeZone": tz or "UTC"}
        changed_fields.append("fim")
    if location is not None:
        changes["location"] = {"displayName": location}
        changed_fields.append("local")
    if body is not None:
        changes["body"] = {"contentType": body_type, "content": body}
        changed_fields.append("corpo")
    if attendees is not None:
        changes["attendees"] = [
            {"emailAddress": {"address": a}, "type": "required"} for a in attendees
        ]
        changed_fields.append("participantes")

    if not changes:
        return {
            "status": "error",
            "message": "Indique pelo menos um campo a alterar.",
        }

    notify = attendees if attendees is not None else (
        [a["email"] for a in event.get("attendees", []) if a.get("email")]
    )
    n = len(notify)
    summary = (
        f"Editar evento {target_event_id}: {', '.join(changed_fields)}. "
        f"Notifica {n} participante(s) (domínios: {', '.join(_domains(notify)) or 'n/d'})."
    )

    prepared = approval.prepare(
        subject=subject,
        account_id=account.account_id,
        operation="calendar.update",
        payload={
            "target_event_id": target_event_id,
            "changes": changes,
            "scope": scope,
            "recipients_count": n,
            "event_subject": event.get("subject"),
        },
        summary=summary,
    )
    prepared["recipients_count"] = n
    return prepared


async def run_calendar_update_confirm(
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
    """US-2.4 — Confirma a edição preparada (token fresco); audita `calendar.update`."""
    async def executor(operation: str, payload: dict) -> dict:
        account, updated = await call_graph(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            op=lambda token: graph_client.update_event(
                token, payload["target_event_id"], changes=payload["changes"],
            ),
            account_id=account_id, clock=clock,
        )
        log_audit(
            audit_logger, action="calendar.update", subject=subject,
            account_id=account.account_id, target=payload["target_event_id"],
            outcome="success", recipients_count=payload.get("recipients_count"),
            extra={
                "subject_hash": subject_hash(payload.get("event_subject") or ""),
                "scope": payload.get("scope"),
            },
        )
        return {
            "operation": operation,
            "event_id": (updated or {}).get("id") or payload["target_event_id"],
            "message": "Evento atualizado.",
        }

    return await _confirm(approval, subject=subject, token=confirmation_token,
                          executor=executor)


async def run_calendar_cancel_prepare(
    subject: str | None,
    *,
    mapping: IdentityMapping,
    plane_b: PlaneB,
    graph_client: GraphClient,
    store: TokenStore,
    approval: ApprovalEngine,
    event_id: str,
    comment: str = "",
    scope: str | None = None,
    account_id: str | None = None,
    clock: Callable[[], datetime] = _utcnow,
) -> dict:
    """US-2.5 — Prepara cancelar (NÃO cancela). Só o organizador pode cancelar (senão, orienta
    para decline). Recorrente sem `scope` -> clarification. D3: resumo declara notificação."""
    if not event_id:
        return {"status": "error", "message": "Indique o event_id a cancelar."}

    try:
        account, _ = await resolve_access_token(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            account_id=account_id, clock=clock,
        )
        tz = await _resolve_tz(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            graph_client=graph_client, account_id=account_id, clock=clock,
        )
        _, event = await call_graph(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            op=lambda token: graph_client.get_event(token, event_id, prefer_timezone=tz),
            account_id=account_id, clock=clock,
        )
        own = await _own_email(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            graph_client=graph_client, account_id=account_id, clock=clock,
        )
    except ReauthRequired as exc:
        return reauth_response(str(exc))

    organizer = event.get("organizer")
    if own and organizer and organizer.lower() != own.lower():
        return {
            "status": "error",
            "message": (
                "Só o organizador pode cancelar este evento. Para deixar de participar, "
                "use calendar_respond com decline."
            ),
        }

    if event.get("isRecurring") and scope is None:
        return _recurrence_clarification("Cancelar")

    target_event_id = event_id
    if scope == "series" and event.get("seriesMasterId"):
        target_event_id = event["seriesMasterId"]

    notify = [a["email"] for a in event.get("attendees", []) if a.get("email")]
    n = len(notify)
    summary = (
        f"Cancelar evento {target_event_id} ('{event.get('subject') or '(sem assunto)'}'). "
        f"Notifica {n} participante(s) (domínios: {', '.join(_domains(notify)) or 'n/d'}) "
        "do cancelamento. Alto impacto."
    )

    prepared = approval.prepare(
        subject=subject,
        account_id=account.account_id,
        operation="calendar.cancel",
        payload={
            "target_event_id": target_event_id,
            "comment": comment,
            "scope": scope,
            "recipients_count": n,
            "event_subject": event.get("subject"),
        },
        summary=summary,
    )
    prepared["recipients_count"] = n
    return prepared


async def run_calendar_cancel_confirm(
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
    """US-2.5 — Confirma o cancelamento (token fresco); audita `calendar.cancel`."""
    async def executor(operation: str, payload: dict) -> dict:
        account, _ = await call_graph(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            op=lambda token: graph_client.cancel_event(
                token, payload["target_event_id"], comment=payload.get("comment", ""),
            ),
            account_id=account_id, clock=clock,
        )
        log_audit(
            audit_logger, action="calendar.cancel", subject=subject,
            account_id=account.account_id, target=payload["target_event_id"],
            outcome="success", recipients_count=payload.get("recipients_count"),
            extra={
                "subject_hash": subject_hash(payload.get("event_subject") or ""),
                "scope": payload.get("scope"),
            },
        )
        return {
            "operation": operation,
            "event_id": payload["target_event_id"],
            "message": "Evento cancelado.",
        }

    return await _confirm(approval, subject=subject, token=confirmation_token,
                          executor=executor)


async def run_calendar_respond_prepare(
    subject: str | None,
    *,
    mapping: IdentityMapping,
    plane_b: PlaneB,
    graph_client: GraphClient,
    store: TokenStore,
    approval: ApprovalEngine,
    event_id: str,
    response: str,
    comment: str = "",
    account_id: str | None = None,
    clock: Callable[[], datetime] = _utcnow,
) -> dict:
    """US-2.6 — Prepara responder a um convite (NÃO responde). D7: lê o estado atual e declara
    a mudança; bloqueia se o subject for organizador. `response` in {accept, decline, tentative}."""
    if not event_id:
        return {"status": "error", "message": "Indique o event_id do convite."}
    if response not in _VALID_RESPONSES:
        return {
            "status": "error",
            "message": "Resposta inválida. Use accept, decline ou tentative.",
        }

    try:
        account, _ = await resolve_access_token(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            account_id=account_id, clock=clock,
        )
        _, event = await call_graph(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            op=lambda token: graph_client.get_event(token, event_id),
            account_id=account_id, clock=clock,
        )
        own = await _own_email(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            graph_client=graph_client, account_id=account_id, clock=clock,
        )
    except ReauthRequired as exc:
        return reauth_response(str(exc))

    organizer = event.get("organizer")
    if own and organizer and organizer.lower() == own.lower():
        return {
            "status": "error",
            "message": (
                "É o organizador deste evento; não pode responder ao próprio convite."
            ),
        }

    previous = event.get("responseStatus") or "none"
    previous_pt = _RESPONSE_PT.get(previous, previous)
    graph_response = _VALID_RESPONSES[response]
    new_pt = _NEW_RESPONSE_PT[graph_response]

    if previous_pt.lower() == new_pt.lower():
        summary = (
            f"Já está como {new_pt}; vai reconfirmar e notificar o organizador."
        )
    else:
        summary = (
            f"Já tinha {previous_pt}; vai mudar para {new_pt} e notificar o organizador."
        )

    prepared = approval.prepare(
        subject=subject,
        account_id=account.account_id,
        operation="calendar.respond",
        payload={
            "event_id": event_id,
            "response": graph_response,
            "comment": comment,
            "previous": previous,
            "event_subject": event.get("subject"),
        },
        summary=summary,
    )
    return prepared


async def run_calendar_respond_confirm(
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
    """US-2.6 — Confirma a resposta ao convite (token fresco); audita `calendar.respond`."""
    async def executor(operation: str, payload: dict) -> dict:
        account, _ = await call_graph(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            op=lambda token: graph_client.respond_event(
                token, payload["event_id"], response=payload["response"],
                comment=payload.get("comment", ""), send_response=True,
            ),
            account_id=account_id, clock=clock,
        )
        log_audit(
            audit_logger, action="calendar.respond", subject=subject,
            account_id=account.account_id, target=payload["event_id"],
            outcome="success",
            extra={
                "subject_hash": subject_hash(payload.get("event_subject") or ""),
                "response": payload["response"],
                "previous": payload.get("previous"),
            },
        )
        return {
            "operation": operation,
            "event_id": payload["event_id"],
            "response": payload["response"],
            "message": "Resposta ao convite enviada.",
        }

    return await _confirm(approval, subject=subject, token=confirmation_token,
                          executor=executor)


async def _confirm(approval, *, subject, token, executor) -> dict:
    """Adaptador comum (idêntico ao email): traduz erros do ApprovalEngine e a reauth do
    executor em respostas amigáveis. Em ReauthRequired o token NÃO é consumido (repetível)."""
    try:
        return await approval.confirm(subject=subject, token=token, executor=executor)
    except ReauthRequired as exc:
        return reauth_response(str(exc))
    except ConfirmationNotFound as exc:
        return {"status": "error", "message": str(exc)}
    except ConfirmationExpired as exc:
        return {"status": "expired", "message": str(exc)}
