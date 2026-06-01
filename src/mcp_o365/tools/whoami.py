"""T11 — Tool `whoami` (read-only) e a sua orquestração.

`whoami` exerce o caminho dual-plane completo: subject (Plano A) -> conta Graph (Plano B)
-> refresh se necessário -> `GET /me`. É a prova end-to-end de que o servidor funciona.
A lógica vive em `run_whoami`, independente do transporte MCP, para ser testável com
Graph/Entra mockados.

Em `invalid_grant` (refresh rejeitado, p.ex. pela Conditional Access) NÃO rebenta: marca a
sessão como expirada e devolve uma mensagem de re-login — a reautenticação graciosa exigida
pela v1.1 §2.2.
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

logger = logging.getLogger("mcp_o365.tools.whoami")

# Margem antes da expiração: renova proativamente para evitar 401 a meio.
_REFRESH_SKEW_SECONDS = 60


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
    if not subject:
        return {"status": "reauth_required", "message": "Sessão não autenticada. Inicie sessão."}

    account = mapping.select_account(subject, account_id)
    if account is None:
        return {
            "status": "reauth_required",
            "message": "Nenhuma conta Office 365 ligada. Inicie sessão para ligar a sua conta.",
        }

    access_token = account.access_token
    now = clock()
    needs_refresh = (
        access_token is None
        or account.expires_at is None
        or (account.expires_at.timestamp() - now.timestamp()) < _REFRESH_SKEW_SECONDS
    )

    if needs_refresh:
        if not account.refresh_token:
            mapping.mark_expired(subject, account.account_id)
            return {
                "status": "reauth_required",
                "message": "Sessão expirada. Volte a iniciar sessão.",
            }
        try:
            refreshed = plane_b.refresh(
                refresh_token=account.refresh_token,
                scopes=account.scopes or None,
                subject_for_log=subject,
                account_id_for_log=account.account_id,
            )
        except ReauthRequired:
            # Inclui InvalidGrant — o sinal típico de bloqueio pela Conditional Access.
            mapping.mark_expired(subject, account.account_id)
            return {
                "status": "reauth_required",
                "message": "A sua sessão expirou ou foi revogada. Volte a iniciar sessão no Claude.",
            }
        access_token = refreshed.access_token
        store.update_account_tokens(
            subject=subject,
            account_id=account.account_id,
            access_token=refreshed.access_token,
            refresh_token=refreshed.refresh_token,
            expires_at=refreshed.expires_at,
        )

    identity = await graph_client.me(access_token)
    return {
        "status": "ok",
        "account_id": account.account_id,
        "id": identity["id"],
        "displayName": identity["displayName"],
        "userPrincipalName": identity["userPrincipalName"],
    }
