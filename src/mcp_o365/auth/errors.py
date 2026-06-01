"""T6 — Exceções de autenticação tipadas.

Permitem que a tool/serviço reaja graciosamente. `ReauthRequired` é o sinal que faz o
assistente pedir novo login dentro do Claude em vez de falhar silenciosamente (v1.1 §2.2).
"""

from __future__ import annotations


class AuthError(Exception):
    """Base de todos os erros de autenticação."""


class ReauthRequired(AuthError):
    """O utilizador tem de voltar a autenticar-se (refresh token expirado/revogado)."""


class InvalidGrant(ReauthRequired):
    """O Entra devolveu `invalid_grant` — caso particular de reauth necessária.

    É também o sinal central do bloqueador de Conditional Access: um refresh silencioso
    rejeitado pela CA chega normalmente como `invalid_grant`/`interaction_required`.
    """


class ConsentRequired(AuthError):
    """É preciso consentimento (admin/utilizador) para os scopes pedidos."""


class UpstreamAuthError(AuthError):
    """Erro inesperado do lado do Entra/Graph que não mapeia para os casos acima."""
