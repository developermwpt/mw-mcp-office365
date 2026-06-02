"""T11 — Tool `whoami` (read-only) e a sua orquestração.

`whoami` exerce o caminho dual-plane completo: subject (Plano A) -> conta Graph (Plano B)
-> refresh se necessário -> `GET /me`. É a prova end-to-end de que o servidor funciona.
A lógica vive em `run_whoami`, independente do transporte MCP, para ser testável com
Graph/Entra mockados.

A resolução de conta + refresh foi extraída para `_session.resolve_access_token` (usada
por todas as tools). Em `invalid_grant` (refresh rejeitado, p.ex. pela Conditional Access)
NÃO rebenta: o helper levanta `ReauthRequired` e aqui converte-se para a mensagem de
re-login — a reautenticação graciosa exigida pela v1.1 §2.2.
"""

from __future__ import annotations

import logging
from collections.abc import Callable
from datetime import datetime, timezone

from ..auth.errors import ReauthRequired
from ..auth.plane_b import PlaneB
from ..graph.client import GraphClient
from ..identity.mapping import IdentityMapping
from ..storage.token_store import TokenStore
from ._session import reauth_response, resolve_access_token

logger = logging.getLogger("mcp_o365.tools.whoami")


def _utcnow() -> datetime:
    return datetime.now(timezone.utc)


async def run_whoami(
    subject: str | None,
    *,
    mapping: IdentityMapping,
    plane_b: PlaneB,
    graph_client: GraphClient,
    store: TokenStore,
    account_id: str | None = None,
    clock: Callable[[], datetime] = _utcnow,
) -> dict:
    """Resolve a identidade do utilizador via Graph. Devolve um dict serializável."""
    try:
        account, access_token = await resolve_access_token(
            subject,
            mapping=mapping,
            plane_b=plane_b,
            store=store,
            account_id=account_id,
            clock=clock,
        )
    except ReauthRequired as exc:
        return reauth_response(str(exc))

    identity = await graph_client.me(access_token)
    return {
        "status": "ok",
        "account_id": account.account_id,
        "id": identity["id"],
        "displayName": identity["displayName"],
        "userPrincipalName": identity["userPrincipalName"],
    }
