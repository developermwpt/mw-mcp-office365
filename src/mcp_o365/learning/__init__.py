"""Fase Aprendizagem (US-L.x) — módulo de aprendizagem do comportamento de email.

Aprende, a partir de METADADOS de ações confirmadas pelo utilizador (mover, responder,
eliminar, arquivar, enviar), padrões reutilizáveis e, mais tarde, gera **recomendações**
de ação que o utilizador apenas tem de confirmar — sempre através do fluxo de aprovação
em duas fases já existente (`*_prepare` -> `confirmation_token` -> `*_confirm`).

Princípios (ver `docs/fase-aprendizagem/analise-funcional-aprendizagem.md`):
- **Opt-in explícito** e **só-metadados**: nunca o corpo nem PII em claro.
- **Isolamento estrito por `subject`**.
- **Explicável e auditável**: cada recomendação traz um `rationale` em linguagem natural.
- **Anti prompt injection**: as features derivam de metadados, NÃO de instruções do corpo.
- **Nunca auto-executar**: recomendar != executar.
"""

from __future__ import annotations

from .features import EmailSignature, extract_signature
from .recommender import Recommendation, Recommender

__all__ = [
    "EmailSignature",
    "Recommendation",
    "Recommender",
    "extract_signature",
]
