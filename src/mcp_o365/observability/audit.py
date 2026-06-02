"""[AC-AUDIT] (v1.1 §1.2) — Auditoria de operações de escrita.

Regista, em formato JSON estruturado, cada operação de escrita (envio, resposta, reencaminho,
mover, eliminar). Por privacidade (RGPD §7) regista APENAS metadados: nunca o corpo do email
nem endereços em claro — quando muito a contagem de destinatários e, no máximo, domínios. O
utilizador é identificado por `subject_hash` (pseudonimização), reutilizando o helper do
`logging_setup`.

Retenção: a v1.1 exige 12 meses de retenção do registo de auditoria. Isso é responsabilidade
de operações (rotação/arquivo dos logs estruturados) — este módulo NÃO implementa purga.
"""

from __future__ import annotations

import logging

from ..logging_setup import subject_hash


def log_audit(
    logger: logging.Logger,
    *,
    action: str,
    subject: str,
    account_id: str | None = None,
    target: str | None = None,
    outcome: str,
    recipients_count: int | None = None,
    extra: dict | None = None,
) -> None:
    """Emite um evento de auditoria (`event=audit`) só com metadados, sem PII.

    `action` é a operação (ex.: `email.send`, `email.delete`). `target` identifica o recurso
    afetado (ex.: id de mensagem/pasta) sem revelar conteúdo. `outcome` é `success`/`error`.
    `recipients_count` substitui a lista de endereços. `extra` permite metadados adicionais
    seguros (ex.: `permanent: true`), mas o chamador é responsável por não incluir PII.
    """
    fields: dict[str, object] = {
        "event": "audit",
        "action": action,
        "subject_hash": subject_hash(subject),
        "account_id": account_id,
        "target": target,
        "outcome": outcome,
    }
    if recipients_count is not None:
        fields["recipients_count"] = recipients_count
    if extra:
        fields.update(extra)
    logger.info("audit", extra={"fields": fields})
