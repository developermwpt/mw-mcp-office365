"""Helper de sessão partilhado pelas tools.

Extrai a lógica de resolução de conta + refresh proativo que antes vivia apenas em
`run_whoami`, para que TODAS as tools (whoami, email, ...) percorram exatamente o mesmo
caminho dual-plane: subject (Plano A) -> conta Graph (Plano B) -> refresh se necessário.

A reautenticação graciosa (v1.1 §2.2) é sinalizada por `ReauthRequired`: o chamador
apanha-a e converte para o dict `reauth_required` com `reauth_response`. Aqui nunca se
devolve o dict diretamente — levanta-se a exceção tipada para manter este helper
componível e testável.

O relógio (`clock`) é injetado para testes determinísticos.
"""

from __future__ import annotations

from collections.abc import Awaitable, Callable
from datetime import datetime, timezone
from typing import Any

from ..auth.errors import ReauthRequired, UpstreamAuthError
from ..auth.plane_b import PlaneB
from ..identity.mapping import IdentityMapping
from ..identity.models import LinkedAccount
from ..storage.token_store import TokenStore

# Margem antes da expiração: renova proativamente para evitar 401 a meio (igual ao whoami).
_REFRESH_SKEW_SECONDS = 60


def _utcnow() -> datetime:
    return datetime.now(timezone.utc)


def reauth_response(message: str) -> dict:
    """Resposta padrão de reautenticação graciosa para o utilizador."""
    return {"status": "reauth_required", "message": message}


async def resolve_access_token(
    subject: str | None,
    *,
    mapping: IdentityMapping,
    plane_b: PlaneB,
    store: TokenStore,
    account_id: str | None = None,
    force_refresh: bool = False,
    clock: Callable[[], datetime] = _utcnow,
) -> tuple[LinkedAccount, str]:
    """Resolve a conta e devolve `(conta, access_token)` pronto a usar.

    Faz refresh proativo se o token expira dentro de `_REFRESH_SKEW_SECONDS` (ou se
    `force_refresh=True`, usado quando o Graph rejeita um token aparentemente válido) e
    persiste os tokens renovados (cifrados) via `store.update_account_tokens`. Levanta
    `ReauthRequired` (com mensagem adequada ao utilizador) quando: o subject está vazio, não
    há conta ligada, falta refresh token, ou o refresh é rejeitado (p.ex. `invalid_grant`).
    """
    if not subject:
        raise ReauthRequired("Sessão não autenticada. Inicie sessão.")

    account = mapping.select_account(subject, account_id)
    if account is None:
        raise ReauthRequired(
            "Nenhuma conta Office 365 ligada. Inicie sessão para ligar a sua conta."
        )

    access_token = account.access_token
    now = clock()
    needs_refresh = (
        force_refresh
        or access_token is None
        or account.expires_at is None
        or (account.expires_at.timestamp() - now.timestamp()) < _REFRESH_SKEW_SECONDS
    )

    if needs_refresh:
        if not account.refresh_token:
            mapping.mark_expired(subject, account.account_id)
            raise ReauthRequired("Sessão expirada. Volte a iniciar sessão.")
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
            raise ReauthRequired(
                "A sua sessão expirou ou foi revogada. Volte a iniciar sessão no Claude."
            ) from None
        access_token = refreshed.access_token
        store.update_account_tokens(
            subject=subject,
            account_id=account.account_id,
            access_token=refreshed.access_token,
            refresh_token=refreshed.refresh_token,
            expires_at=refreshed.expires_at,
        )

    return account, access_token


async def call_graph(
    subject: str | None,
    *,
    mapping: IdentityMapping,
    plane_b: PlaneB,
    store: TokenStore,
    op: Callable[[str], Awaitable[Any]],
    account_id: str | None = None,
    clock: Callable[[], datetime] = _utcnow,
) -> tuple[LinkedAccount, Any]:
    """Resolve `(conta, token)` e corre `op(token)` com resiliência a auth do Graph.

    Se o Graph rejeitar o token (`UpstreamAuthError`, tipicamente 401/403 — token expirado,
    rotacionado, ou scopes acabados de mudar) apesar de parecer válido, **força um refresh e
    repete uma vez**. Se mesmo assim falhar, marca a conta como expirada e levanta
    `ReauthRequired` — o chamador converte em `reauth_required` (nunca um erro cru ao
    utilizador). Devolve `(conta, resultado_de_op)`.
    """
    account, token = await resolve_access_token(
        subject, mapping=mapping, plane_b=plane_b, store=store,
        account_id=account_id, clock=clock,
    )
    try:
        return account, await op(token)
    except UpstreamAuthError:
        # Token recusado pelo Graph — força refresh e tenta de novo (uma vez).
        account, token = await resolve_access_token(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            account_id=account_id, force_refresh=True, clock=clock,
        )
        try:
            return account, await op(token)
        except UpstreamAuthError as exc:
            mapping.mark_expired(subject, account.account_id)
            raise ReauthRequired(
                "A sua sessão expirou ou foi revogada. Volte a iniciar sessão no Claude."
            ) from exc
