"""T6 — Plano B: client confidencial contra o Entra ID (via msal).

Núcleo do bloqueador de Conditional Access: é aqui que vive o **refresh** do token Graph.
Gera a URL de autorização, troca o código por tokens, e renova por refresh token,
normalizando os erros do Entra para as exceções tipadas de `errors.py`.

A aplicação msal é injetada por um factory (`msal_app_factory`) para ser substituível por
um fake nos testes — nenhuma rede real é exigida.
"""

from __future__ import annotations

import logging
from collections.abc import Callable
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from typing import Any

from ..config import Settings
from ..logging_setup import log_refresh_failure
from .errors import ConsentRequired, InvalidGrant, UpstreamAuthError

logger = logging.getLogger("mcp_o365.auth.plane_b")

# Scopes que o MSAL adiciona automaticamente — não podem ser passados explicitamente.
_MSAL_RESERVED_SCOPES: frozenset[str] = frozenset({"offline_access", "openid", "profile", "email"})


def _graph_scopes(scopes: list[str] | None) -> list[str]:
    """Remove os scopes reservados pelo MSAL antes de os passar às suas APIs."""
    return [s for s in (scopes or []) if s not in _MSAL_RESERVED_SCOPES]


# Erros do Entra que significam "é preciso interação do utilizador".
_REAUTH_ERRORS = {"invalid_grant", "interaction_required", "login_required"}
_CONSENT_ERRORS = {"consent_required"}


def _utcnow() -> datetime:
    return datetime.now(timezone.utc)


@dataclass
class TokenResult:
    """Resultado normalizado de uma operação de token do Plano B."""

    access_token: str
    refresh_token: str | None
    expires_at: datetime | None
    scopes: list[str]
    home_account_id: str | None
    username: str | None
    tenant_id: str | None
    claims: dict[str, Any]


class PlaneB:
    """Orquestra o fluxo OAuth Authorization Code contra o Entra ID."""

    def __init__(
        self,
        config: Settings,
        msal_app_factory: Callable[[], Any] | None = None,
        clock: Callable[[], datetime] = _utcnow,
    ) -> None:
        self._config = config
        self._clock = clock
        self._factory = msal_app_factory or self._default_factory

    def _default_factory(self) -> Any:
        # Importação adiada: msal só é necessário em runtime real, não nos testes.
        import msal

        return msal.ConfidentialClientApplication(
            client_id=self._config.entra_client_id,
            authority=self._config.entra_authority,
            client_credential=self._config.entra_client_secret.get_secret_value(),
        )

    def build_authorization_url(
        self, *, state: str, redirect_uri: str, scopes: list[str] | None = None
    ) -> str:
        """URL de autorização do Entra para onde redirecionar o browser do utilizador."""
        app = self._factory()
        effective = _graph_scopes(scopes or self._config.graph_scopes)
        return app.get_authorization_request_url(
            effective,
            state=state,
            redirect_uri=redirect_uri,
        )

    def exchange_code(
        self, *, code: str, redirect_uri: str, scopes: list[str] | None = None
    ) -> TokenResult:
        """Troca o authorization code do Entra por tokens Graph."""
        app = self._factory()
        effective = _graph_scopes(scopes or self._config.graph_scopes)
        result = app.acquire_token_by_authorization_code(
            code,
            scopes=effective,
            redirect_uri=redirect_uri,
        )
        return self._normalize(result, subject_for_log=None)

    def refresh(
        self,
        *,
        refresh_token: str,
        scopes: list[str] | None = None,
        subject_for_log: str | None = None,
        account_id_for_log: str | None = None,
    ) -> TokenResult:
        """Renova o token Graph. Em `invalid_grant`, regista `refresh_failure` e levanta."""
        app = self._factory()
        effective = _graph_scopes(scopes or self._config.graph_scopes)
        result = app.acquire_token_by_refresh_token(
            refresh_token, scopes=effective
        )
        return self._normalize(
            result,
            subject_for_log=subject_for_log,
            account_id_for_log=account_id_for_log,
            is_refresh=True,
        )

    def _normalize(
        self,
        result: dict[str, Any],
        *,
        subject_for_log: str | None,
        account_id_for_log: str | None = None,
        is_refresh: bool = False,
    ) -> TokenResult:
        error = result.get("error")
        if error:
            reason = result.get("error_description", error)
            if is_refresh and subject_for_log:
                # Sinal-chave para o diagnóstico do bloqueador de CA.
                log_refresh_failure(
                    logger,
                    subject=subject_for_log,
                    account_id=account_id_for_log,
                    reason=error,
                )
            if error in _REAUTH_ERRORS:
                raise InvalidGrant(reason)
            if error in _CONSENT_ERRORS:
                raise ConsentRequired(reason)
            raise UpstreamAuthError(f"{error}: {reason}")

        if "access_token" not in result:
            raise UpstreamAuthError("Resposta do Entra sem access_token.")

        expires_in = result.get("expires_in")
        expires_at = (
            self._clock() + timedelta(seconds=int(expires_in)) if expires_in else None
        )
        claims = result.get("id_token_claims", {}) or {}
        scope_value = result.get("scope")
        scopes = scope_value.split() if isinstance(scope_value, str) else (scope_value or [])
        return TokenResult(
            access_token=result["access_token"],
            refresh_token=result.get("refresh_token"),
            expires_at=expires_at,
            scopes=scopes,
            home_account_id=claims.get("oid") or claims.get("sub"),
            username=claims.get("preferred_username") or claims.get("upn"),
            tenant_id=claims.get("tid"),
            claims=claims,
        )
