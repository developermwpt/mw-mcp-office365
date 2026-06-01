"""T2 — Logging estruturado (JSON), sem PII.

Cada linha de log é um objeto JSON. Um `request_id` propaga-se por `contextvars`.
Há um evento dedicado `refresh_failure` — o sinal-chave para diagnosticar o bloqueador
de Conditional Access (v1.1 §2.4). Nunca se registam tokens, UPN ou outros dados pessoais
em claro; o utilizador é identificado por um hash truncado do `subject`.
"""

from __future__ import annotations

import contextvars
import hashlib
import json
import logging
import sys
from datetime import datetime, timezone

_request_id: contextvars.ContextVar[str | None] = contextvars.ContextVar(
    "request_id", default=None
)


def set_request_id(value: str | None) -> None:
    _request_id.set(value)


def subject_hash(subject: str) -> str:
    """Hash estável e truncado do subject — pseudonimização para logs (RGPD §7)."""
    return hashlib.sha256(subject.encode("utf-8")).hexdigest()[:16]


class JsonFormatter(logging.Formatter):
    """Formata cada registo como uma linha JSON com metadados de correlação."""

    def format(self, record: logging.LogRecord) -> str:
        payload: dict[str, object] = {
            "ts": datetime.now(timezone.utc).isoformat(),
            "level": record.levelname,
            "logger": record.name,
            "msg": record.getMessage(),
        }
        rid = _request_id.get()
        if rid:
            payload["request_id"] = rid
        # Campos extra colocados via `logger.info(..., extra={"fields": {...}})`.
        fields = getattr(record, "fields", None)
        if isinstance(fields, dict):
            payload.update(fields)
        if record.exc_info:
            payload["exc"] = self.formatException(record.exc_info)
        return json.dumps(payload, ensure_ascii=False, sort_keys=True)


def setup_logging(level: str = "INFO") -> None:
    """Configura o root logger para emitir JSON em stdout. Idempotente."""
    handler = logging.StreamHandler(stream=sys.stdout)
    handler.setFormatter(JsonFormatter())
    root = logging.getLogger()
    root.handlers.clear()
    root.addHandler(handler)
    root.setLevel(level.upper())


def log_refresh_failure(
    logger: logging.Logger,
    *,
    subject: str,
    account_id: str | None,
    reason: str,
) -> None:
    """Emite o evento `refresh_failure` com schema fixo e sem PII.

    É o sinal observável que distingue uma CA a bloquear o refresh silencioso de uma
    simples expiração de refresh token.
    """
    logger.warning(
        "refresh falhou",
        extra={
            "fields": {
                "event": "refresh_failure",
                "subject_hash": subject_hash(subject),
                "account_id": account_id,
                "reason": reason,
            }
        },
    )
