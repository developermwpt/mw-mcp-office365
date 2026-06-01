"""T5 — Mapeamento de identidade (o elo crítico do dual-plane).

Dado o `subject` do token do Plano A, resolve a `GraphSession` correta do Plano B, com
isolamento estrito por utilizador. Cria/atualiza sessões e contas, seleciona a conta a
usar (default ou específica) e marca sessões/contas como expiradas para a reautenticação
graciosa (v1.1 §2.2).
"""

from __future__ import annotations

from collections.abc import Callable
from datetime import datetime, timezone

from ..storage.token_store import TokenStore
from .models import GraphSession, LinkedAccount


def _utcnow() -> datetime:
    return datetime.now(timezone.utc)


class IdentityMapping:
    """Resolve subjects (Plano A) em sessões Graph (Plano B)."""

    def __init__(self, store: TokenStore, clock: Callable[[], datetime] = _utcnow) -> None:
        self._store = store
        self._clock = clock

    def get_session(self, subject: str) -> GraphSession | None:
        """Devolve a sessão ativa do subject, ou `None` se não existir (-> reauth)."""
        row = self._store.get_active_session(subject)
        if row is None:
            return None
        accounts = [LinkedAccount.from_row(a) for a in self._store.list_accounts(subject)]
        return GraphSession(
            session_id=row["session_id"], subject=subject, accounts=accounts
        )

    def ensure_session(self, subject: str) -> str:
        """Garante uma sessão ativa para o subject; devolve o `session_id`."""
        row = self._store.get_active_session(subject)
        if row is not None:
            return row["session_id"]
        return self._store.create_session(subject)

    def link_account(
        self,
        *,
        subject: str,
        access_token: str,
        refresh_token: str | None,
        expires_at: datetime | None,
        tenant_id: str | None = None,
        home_account_id: str | None = None,
        username: str | None = None,
        scopes: list[str] | None = None,
        is_default: bool = True,
    ) -> str:
        """Cria/atualiza uma conta O365 na sessão do subject. Devolve o `account_id`."""
        session_id = self.ensure_session(subject)
        return self._store.upsert_account(
            subject=subject,
            session_id=session_id,
            access_token=access_token,
            refresh_token=refresh_token,
            expires_at=expires_at,
            tenant_id=tenant_id,
            home_account_id=home_account_id,
            username=username,
            scopes=scopes,
            is_default=is_default,
        )

    def select_account(
        self, subject: str, account_id: str | None = None
    ) -> LinkedAccount | None:
        """Seleciona a conta a usar: a indicada, ou a default. `None` se nenhuma ativa."""
        row = (
            self._store.get_account(subject, account_id)
            if account_id
            else self._store.get_default_account(subject)
        )
        return LinkedAccount.from_row(row) if row else None

    def mark_expired(self, subject: str, account_id: str | None = None) -> None:
        """Marca a conta (ou a sessão inteira) como expirada — força reauth graciosa."""
        if account_id:
            self._store.mark_account_expired(subject, account_id)
        else:
            self._store.mark_session_expired(subject)
