"""T5 — Modelos de identidade.

Multi-conta desde o início: uma `GraphSession` (do Plano A) agrega N `LinkedAccount`
(contas O365 do Plano B). Adicionar multi-conta mais tarde implicaria reescrita; por isso
o modelo nasce assim (v1.1 §2.2).
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime


@dataclass(frozen=True)
class McpPrincipal:
    """Quem o token do Plano A diz que o utilizador é."""

    subject: str


@dataclass
class LinkedAccount:
    """Uma conta O365 ligada (Plano B), com os tokens Graph delegados."""

    account_id: str
    subject: str
    session_id: str
    tenant_id: str | None = None
    home_account_id: str | None = None
    username: str | None = None
    scopes: list[str] = field(default_factory=list)
    access_token: str | None = None
    refresh_token: str | None = None
    expires_at: datetime | None = None
    is_default: bool = False
    status: str = "active"

    @classmethod
    def from_row(cls, row: dict) -> LinkedAccount:
        return cls(
            account_id=row["account_id"],
            subject=row["subject"],
            session_id=row["session_id"],
            tenant_id=row.get("tenant_id"),
            home_account_id=row.get("home_account_id"),
            username=row.get("username"),
            scopes=row.get("scopes", []),
            access_token=row.get("access_token"),
            refresh_token=row.get("refresh_token"),
            expires_at=row.get("expires_at"),
            is_default=row.get("is_default", False),
            status=row.get("status", "active"),
        )

    def is_expired(self, now: datetime) -> bool:
        return self.expires_at is not None and self.expires_at <= now


@dataclass
class GraphSession:
    """Sessão de um principal do Plano A, com as suas contas O365."""

    session_id: str
    subject: str
    accounts: list[LinkedAccount] = field(default_factory=list)

    @property
    def default_account(self) -> LinkedAccount | None:
        for acc in self.accounts:
            if acc.is_default:
                return acc
        return self.accounts[0] if self.accounts else None
