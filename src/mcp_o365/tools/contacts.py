"""Módulo 5 — Resolução de destinatários por nome (US-5.x).

`run_resolve_recipient` recebe um NOME (ex.: "vera") e devolve candidatos
`{display_name, email, source}` ordenados por relevância, juntando o People API
(ranqueado pelo Graph) e os Contactos pessoais. É READ-ONLY: não envia nem agenda nada
— apenas resolve para o utilizador confirmar antes de qualquer `*_prepare`.

Padrão de confirmação (igual ao `email_reply_prepare`): 0 -> `not_found`; 1 -> `ok` com o
candidato; vários -> `needs_clarification` para o utilizador escolher. Nunca escolhe sozinho
quando é ambíguo. Reautenticação graciosa via `call_graph` (401/403 -> reauth).
"""

from __future__ import annotations

import logging
from collections.abc import Callable
from datetime import datetime, timezone

from ..auth.errors import ReauthRequired
from ..auth.plane_b import PlaneB
from ..graph.client import GraphClient
from ..identity.mapping import IdentityMapping
from ..observability.audit import log_audit
from ..storage.token_store import TokenStore
from ._session import call_graph, reauth_response

logger = logging.getLogger("mcp_o365.tools.contacts")
audit_logger = logging.getLogger("mcp_o365.audit")


def _utcnow() -> datetime:
    return datetime.now(timezone.utc)


def _merge_candidates(people: list[dict], contacts: list[dict]) -> list[dict]:
    """Junta People (primeiro, mais relevante) + Contactos, deduplicando por email."""
    out: list[dict] = []
    seen: set[str] = set()
    for cand in [*people, *contacts]:
        email = (cand.get("email") or "").strip()
        if not email:
            continue  # sem email não serve para enviar/agendar
        key = email.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(
            {
                "display_name": cand.get("display_name"),
                "email": email,
                "source": cand.get("source"),
            }
        )
    return out


async def run_resolve_recipient(
    subject: str | None,
    *,
    mapping: IdentityMapping,
    plane_b: PlaneB,
    graph_client: GraphClient,
    store: TokenStore,
    name: str,
    top: int = 10,
    account_id: str | None = None,
    clock: Callable[[], datetime] = _utcnow,
) -> dict:
    """US-5.1 — Resolve um nome em candidatos a destinatário (read-only)."""
    if not name or not name.strip():
        return {"status": "error", "message": "Indique um nome a procurar."}
    query = name.strip()

    async def op(token: str) -> tuple[list[dict], list[dict]]:
        people = await graph_client.search_people(token, query, top=top)
        contacts = await graph_client.search_contacts(token, query, top=top)
        return people, contacts

    try:
        _, (people, contacts) = await call_graph(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            op=op, account_id=account_id, clock=clock,
        )
    except ReauthRequired as exc:
        return reauth_response(str(exc))

    candidates = _merge_candidates(people, contacts)

    log_audit(
        audit_logger, action="contacts.resolve", subject=subject or "",
        outcome="success", extra={"query": query, "count": len(candidates)},
    )

    if not candidates:
        return {
            "status": "not_found",
            "query": query,
            "candidates": [],
            "message": f"Não encontrei ninguém como '{query}' nos seus contactos/pessoas.",
        }
    if len(candidates) == 1:
        return {
            "status": "ok",
            "query": query,
            "candidates": candidates,
            "recipient": candidates[0],
            "message": (
                f"Encontrei {candidates[0]['display_name'] or candidates[0]['email']} "
                f"<{candidates[0]['email']}>. Confirme antes de usar."
            ),
        }
    return {
        "status": "needs_clarification",
        "query": query,
        "candidates": candidates,
        "message": (
            f"Encontrei vários candidatos para '{query}'. Pergunte ao utilizador qual usar "
            "(pelo email) antes de enviar/agendar."
        ),
    }
