"""Motor de aprovação em duas fases (v1.1 §3).

`prepare` persiste a operação pendente (payload cifrado pelo store) e devolve um token de
uso único com TTL. `confirm` valida o token e, só então, executa a operação através de um
`executor` fornecido de fora — o engine NÃO conhece o Microsoft Graph, o que o mantém
testável sem rede.

Idempotência: se um token já consumido (com resultado guardado) for reapresentado, devolve
o resultado anterior SEM re-executar (com a flag `idempotent_replay`). Isto protege contra
retries do cliente que de outra forma duplicariam envios/eliminações.

O relógio (`clock`) é injetado para testes determinísticos.
"""

from __future__ import annotations

import uuid
from collections.abc import Awaitable, Callable
from datetime import datetime, timedelta, timezone

from ..storage.token_store import TokenStore
from .errors import ConfirmationExpired, ConfirmationNotFound

# Executor: recebe (operation, payload) e devolve o dict de resultado da execução real.
Executor = Callable[[str, dict], Awaitable[dict]]


def _utcnow() -> datetime:
    return datetime.now(timezone.utc)


class ApprovalEngine:
    """Orquestra o ciclo prepare -> confirm sobre o `TokenStore`."""

    def __init__(
        self,
        store: TokenStore,
        clock: Callable[[], datetime] = _utcnow,
        ttl_seconds: int = 300,
    ) -> None:
        self._store = store
        self._clock = clock
        self._ttl = ttl_seconds

    def prepare(
        self,
        *,
        subject: str,
        account_id: str | None,
        operation: str,
        payload: dict,
        summary: str,
    ) -> dict:
        """Regista a operação pendente e devolve o token de confirmação + resumo."""
        token = uuid.uuid4().hex
        expires_at = self._clock() + timedelta(seconds=self._ttl)
        self._store.save_pending_operation(
            token=token,
            subject=subject,
            account_id=account_id,
            operation=operation,
            payload=payload,
            summary=summary,
            expires_at=expires_at,
        )
        return {
            "status": "pending_confirmation",
            "operation": operation,
            "summary": summary,
            "confirmation_token": token,
            "expires_at": expires_at.isoformat(),
        }

    async def confirm(
        self,
        *,
        subject: str,
        token: str,
        executor: Executor,
    ) -> dict:
        """Valida o token e executa a operação via `executor`. Idempotente em replay."""
        op = self._store.get_pending_operation(subject, token)
        if op is None:
            raise ConfirmationNotFound(
                "Token de confirmação inválido ou desconhecido."
            )

        # Idempotência: já consumido e com resultado guardado -> devolve sem re-executar.
        if op["consumed_at"] is not None and op["result"] is not None:
            return {"status": "done", "idempotent_replay": True, **op["result"]}

        if self._clock() > op["expires_at"]:
            raise ConfirmationExpired(
                "O pedido de confirmação expirou. Volte a preparar a operação."
            )

        result = await executor(op["operation"], op["payload"])
        self._store.mark_pending_consumed(subject=subject, token=token, result=result)
        return {"status": "done", **result}
