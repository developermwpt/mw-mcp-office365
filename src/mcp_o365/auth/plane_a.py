"""T9 — Plano A: o MCP como OAuth Authorization/Resource Server para o Claude.

# NOTA SDK: o `mcp` 1.27.2 roteia `/authorize`, `/token`, `/register`, `/revoke` e a
# metadata para os métodos deste provider. Implementamos o `OAuthAuthorizationServerProvider`
# e um `TokenVerifier`; o SDK trata do transporte HTTP e da validação de PKCE do lado do
# Claude. A perna do Entra (Plano B) é orquestrada por `complete_entra_callback`, chamada
# pela rota custom `/callback` (server.py) — o SDK não a conhece.

Fluxo:
  Claude --/authorize--> authorize(): cria `state`, guarda a transação, redireciona para o
      Entra (via Plano B).
  Entra --/callback--> complete_entra_callback(): troca o code do Entra por tokens Graph,
      resolve o `subject`, liga a conta (mapping), emite um authorization code do Plano A e
      redireciona de volta para o Claude.
  Claude --/token--> exchange_authorization_code(): emite o access token MCP (opaco),
      ligado ao `subject`.
  Tool ---------> verify_token()/load_access_token(): resolve o token MCP -> subject.
"""

from __future__ import annotations

import logging
import secrets
from collections.abc import Callable
from datetime import datetime, timedelta, timezone

from mcp.server.auth.provider import (
    AccessToken,
    AuthorizationCode,
    AuthorizationParams,
    OAuthAuthorizationServerProvider,
    RefreshToken,
    TokenError,
    TokenVerifier,
    construct_redirect_uri,
)
from mcp.shared.auth import OAuthClientInformationFull, OAuthToken

from ..config import Settings
from ..identity.mapping import IdentityMapping
from ..storage.token_store import TokenStore
from .errors import AuthError
from .plane_b import PlaneB

logger = logging.getLogger("mcp_o365.auth.plane_a")

_AUTH_CODE_TTL = timedelta(minutes=5)
_ACCESS_TOKEN_TTL = timedelta(hours=1)
_REFRESH_TOKEN_TTL = timedelta(days=30)


def _utcnow() -> datetime:
    return datetime.now(timezone.utc)


class MwOAuthProvider(
    OAuthAuthorizationServerProvider[AuthorizationCode, RefreshToken, AccessToken]
):
    """Provider OAuth do Plano A, com a ponte para o Plano B (Entra)."""

    def __init__(
        self,
        *,
        store: TokenStore,
        plane_b: PlaneB,
        mapping: IdentityMapping,
        config: Settings,
        clock: Callable[[], datetime] = _utcnow,
    ) -> None:
        self._store = store
        self._plane_b = plane_b
        self._mapping = mapping
        self._config = config
        self._clock = clock

    # --- DCR (RFC 7591) ---
    async def get_client(self, client_id: str) -> OAuthClientInformationFull | None:
        raw = self._store.get_client(client_id)
        return OAuthClientInformationFull.model_validate_json(raw) if raw else None

    async def register_client(self, client_info: OAuthClientInformationFull) -> None:
        self._store.save_client(client_info.client_id, client_info.model_dump_json())

    # --- /authorize (Plano A) -> redireciona para o Entra (Plano B) ---
    async def authorize(
        self, client: OAuthClientInformationFull, params: AuthorizationParams
    ) -> str:
        state = secrets.token_urlsafe(32)
        self._store.save_transaction(
            state=state,
            client_id=client.client_id,
            client_redirect_uri=str(params.redirect_uri),
            client_code_challenge=params.code_challenge,
            client_state=params.state,
            scopes=params.scopes or self._config.graph_scopes,
        )
        return self._plane_b.build_authorization_url(
            state=state, redirect_uri=self._config.oauth_redirect_uri
        )

    # --- /callback do Entra (rota custom) ---
    def complete_entra_callback(self, *, code: str, state: str) -> str:
        """Conclui o login do Entra e devolve a URL de redireção de volta ao Claude."""
        tx = self._store.pop_transaction(state)
        if tx is None:
            raise AuthError("Transação de autorização desconhecida ou expirada (state).")

        token = self._plane_b.exchange_code(
            code=code,
            redirect_uri=self._config.oauth_redirect_uri,
            scopes=tx["scopes"],
        )
        subject = token.home_account_id or token.username
        if not subject:
            raise AuthError("Não foi possível resolver o subject a partir do Entra.")

        # Liga a conta O365 (Plano B) ao subject (Plano A); guarda tokens cifrados.
        self._mapping.link_account(
            subject=subject,
            access_token=token.access_token,
            refresh_token=token.refresh_token,
            expires_at=token.expires_at,
            tenant_id=token.tenant_id,
            home_account_id=token.home_account_id,
            username=token.username,
            scopes=token.scopes,
            is_default=True,
        )

        # Emite o authorization code do Plano A (ligado ao subject) para o Claude.
        plana_code = secrets.token_urlsafe(32)
        self._store.save_auth_code(
            code=plana_code,
            client_id=tx["client_id"],
            subject=subject,
            code_challenge=tx["client_code_challenge"] or "",
            redirect_uri=tx["client_redirect_uri"],
            redirect_uri_explicit=True,
            scopes=tx["scopes"],
            expires_at=self._clock() + _AUTH_CODE_TTL,
        )
        params = {"code": plana_code}
        if tx["client_state"]:
            params["state"] = tx["client_state"]
        return construct_redirect_uri(tx["client_redirect_uri"], **params)

    # --- authorization code (Plano A) ---
    async def load_authorization_code(
        self, client: OAuthClientInformationFull, authorization_code: str
    ) -> AuthorizationCode | None:
        row = self._store.get_auth_code(authorization_code)
        if row is None or row["client_id"] != client.client_id:
            return None
        return AuthorizationCode(
            code=row["code"],
            scopes=row["scopes"],
            expires_at=row["expires_at"].timestamp(),
            client_id=row["client_id"],
            code_challenge=row["code_challenge"] or "",
            redirect_uri=row["redirect_uri"],
            redirect_uri_provided_explicitly=row["redirect_uri_explicit"],
            subject=row["subject"],
        )

    async def exchange_authorization_code(
        self, client: OAuthClientInformationFull, authorization_code: AuthorizationCode
    ) -> OAuthToken:
        row = self._store.pop_auth_code(authorization_code.code)
        if row is None:
            raise TokenError("invalid_grant", "Authorization code inválido ou já usado.")
        if row["expires_at"] and row["expires_at"] < self._clock():
            raise TokenError("invalid_grant", "Authorization code expirado.")

        subject = row["subject"]
        scopes = row["scopes"]
        access = secrets.token_urlsafe(32)
        refresh = secrets.token_urlsafe(32)
        now = self._clock()
        self._store.save_access_token(
            token=access, client_id=client.client_id, subject=subject,
            scopes=scopes, expires_at=now + _ACCESS_TOKEN_TTL,
        )
        self._store.save_refresh_token(
            token=refresh, client_id=client.client_id, subject=subject,
            scopes=scopes, expires_at=now + _REFRESH_TOKEN_TTL,
        )
        return OAuthToken(
            access_token=access,
            token_type="Bearer",
            expires_in=int(_ACCESS_TOKEN_TTL.total_seconds()),
            scope=" ".join(scopes) if scopes else None,
            refresh_token=refresh,
        )

    # --- refresh token (Plano A) ---
    async def load_refresh_token(
        self, client: OAuthClientInformationFull, refresh_token: str
    ) -> RefreshToken | None:
        row = self._store.get_refresh_token(refresh_token)
        if row is None or row["client_id"] != client.client_id:
            return None
        return RefreshToken(
            token=row["token"],
            client_id=row["client_id"],
            scopes=row["scopes"],
            expires_at=int(row["expires_at"].timestamp()) if row["expires_at"] else None,
            subject=row["subject"],
        )

    async def exchange_refresh_token(
        self,
        client: OAuthClientInformationFull,
        refresh_token: RefreshToken,
        scopes: list[str],
    ) -> OAuthToken:
        # Rotação simples do refresh token do Plano A.
        self._store.delete_refresh_token(refresh_token.token)
        subject = refresh_token.subject or ""
        granted = scopes or refresh_token.scopes
        access = secrets.token_urlsafe(32)
        new_refresh = secrets.token_urlsafe(32)
        now = self._clock()
        self._store.save_access_token(
            token=access, client_id=client.client_id, subject=subject,
            scopes=granted, expires_at=now + _ACCESS_TOKEN_TTL,
        )
        self._store.save_refresh_token(
            token=new_refresh, client_id=client.client_id, subject=subject,
            scopes=granted, expires_at=now + _REFRESH_TOKEN_TTL,
        )
        return OAuthToken(
            access_token=access,
            token_type="Bearer",
            expires_in=int(_ACCESS_TOKEN_TTL.total_seconds()),
            scope=" ".join(granted) if granted else None,
            refresh_token=new_refresh,
        )

    # --- access token (Plano A) ---
    async def load_access_token(self, token: str) -> AccessToken | None:
        return self._resolve_access_token(token)

    async def revoke_token(self, token: AccessToken | RefreshToken) -> None:
        self._store.delete_access_token(token.token)
        self._store.delete_refresh_token(token.token)

    def _resolve_access_token(self, token: str) -> AccessToken | None:
        row = self._store.get_access_token(token)
        if row is None:
            return None
        if row["expires_at"] and row["expires_at"] < self._clock():
            self._store.delete_access_token(token)
            return None
        return AccessToken(
            token=row["token"],
            client_id=row["client_id"],
            scopes=row["scopes"],
            expires_at=int(row["expires_at"].timestamp()) if row["expires_at"] else None,
            subject=row["subject"],
        )


class MwTokenVerifier(TokenVerifier):
    """Verifica o access token MCP (Plano A) recebido em cada request de tool."""

    def __init__(self, provider: MwOAuthProvider) -> None:
        self._provider = provider

    async def verify_token(self, token: str) -> AccessToken | None:
        return self._provider._resolve_access_token(token)
