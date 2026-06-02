"""Exceções tipadas do motor de aprovação em duas fases.

Permitem que as tools distingam "token inexistente/já usado sem resultado" de "token
expirado" e respondam ao utilizador com a mensagem adequada (sem expor stack traces).
"""

from __future__ import annotations


class ApprovalError(Exception):
    """Base de todos os erros do fluxo de aprovação."""


class ConfirmationNotFound(ApprovalError):
    """O token de confirmação não existe (ou não pertence a este subject)."""


class ConfirmationExpired(ApprovalError):
    """O token de confirmação expirou — é preciso voltar a preparar a operação."""
