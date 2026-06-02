"""Aprovação em duas fases (prepare -> confirm) para operações de escrita.

Padrão da v1.1 §3: nenhuma operação de escrita executa sem confirmação explícita do
utilizador. O `prepare` valida, monta e devolve um resumo legível mais um token de uso
único com TTL; o `confirm` só executa se esse token for válido, não estiver expirado nem
já consumido. O token serve também de idempotency key.
"""

from __future__ import annotations

from .engine import ApprovalEngine
from .errors import ApprovalError, ConfirmationExpired, ConfirmationNotFound

__all__ = [
    "ApprovalEngine",
    "ApprovalError",
    "ConfirmationExpired",
    "ConfirmationNotFound",
]
