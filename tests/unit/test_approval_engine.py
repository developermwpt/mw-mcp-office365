"""Unit — ApprovalEngine (v1.1 §3): prepare/confirm, idempotência, TTL, isolamento.

O `executor` é um fake assíncrono que conta quantas vezes foi chamado, para provar que o
replay NÃO re-executa a operação real (proteção contra retries do cliente).
"""

from __future__ import annotations

import pytest

from mcp_o365.approval.engine import ApprovalEngine
from mcp_o365.approval.errors import ConfirmationExpired, ConfirmationNotFound


class CountingExecutor:
    """Executor de teste: regista cada invocação e devolve um resultado fixo."""

    def __init__(self, result: dict | None = None) -> None:
        self.calls: list[tuple[str, dict]] = []
        self._result = result or {"message": "executado"}

    async def __call__(self, operation: str, payload: dict) -> dict:
        self.calls.append((operation, payload))
        return self._result


def _prepare(engine: ApprovalEngine, *, subject: str = "subj-1") -> dict:
    return engine.prepare(
        subject=subject,
        account_id="acc-1",
        operation="email.send",
        payload={"x": 1},
        summary="Enviar email de teste.",
    )


def test_prepare_devolve_token_resumo_e_expires(store, clock):
    engine = ApprovalEngine(store, clock=clock, ttl_seconds=300)
    prepared = _prepare(engine)
    assert prepared["status"] == "pending_confirmation"
    assert prepared["operation"] == "email.send"
    assert prepared["summary"] == "Enviar email de teste."
    assert prepared["confirmation_token"]
    assert prepared["expires_at"]  # ISO string


async def test_confirm_executa_uma_vez_e_devolve_done(store, clock):
    engine = ApprovalEngine(store, clock=clock)
    prepared = _prepare(engine)
    executor = CountingExecutor({"message": "ok"})

    out = await engine.confirm(
        subject="subj-1", token=prepared["confirmation_token"], executor=executor
    )

    assert out["status"] == "done"
    assert out["message"] == "ok"
    assert len(executor.calls) == 1
    assert executor.calls[0][0] == "email.send"


async def test_confirm_idempotente_no_replay(store, clock):
    """Segundo confirm com o mesmo token NÃO re-executa; devolve idempotent_replay."""
    engine = ApprovalEngine(store, clock=clock)
    prepared = _prepare(engine)
    executor = CountingExecutor({"message": "enviado"})
    token = prepared["confirmation_token"]

    first = await engine.confirm(subject="subj-1", token=token, executor=executor)
    second = await engine.confirm(subject="subj-1", token=token, executor=executor)

    assert first["status"] == "done"
    assert "idempotent_replay" not in first
    assert second["status"] == "done"
    assert second["idempotent_replay"] is True
    assert second["message"] == "enviado"
    # O executor só correu UMA vez — o replay não duplica a operação.
    assert len(executor.calls) == 1


async def test_confirm_ttl_expirado_levanta(store, clock):
    engine = ApprovalEngine(store, clock=clock, ttl_seconds=300)
    prepared = _prepare(engine)
    executor = CountingExecutor()

    clock.advance(301)  # para além do TTL

    with pytest.raises(ConfirmationExpired):
        await engine.confirm(
            subject="subj-1", token=prepared["confirmation_token"], executor=executor
        )
    assert executor.calls == []


async def test_confirm_token_desconhecido_levanta(store, clock):
    engine = ApprovalEngine(store, clock=clock)
    executor = CountingExecutor()
    with pytest.raises(ConfirmationNotFound):
        await engine.confirm(
            subject="subj-1", token="token-que-nao-existe", executor=executor
        )
    assert executor.calls == []


async def test_confirm_isolamento_entre_subjects(store, clock):
    """O subject B não vê o token preparado pelo subject A -> ConfirmationNotFound."""
    engine = ApprovalEngine(store, clock=clock)
    prepared = _prepare(engine, subject="subj-A")
    executor = CountingExecutor()

    with pytest.raises(ConfirmationNotFound):
        await engine.confirm(
            subject="subj-B", token=prepared["confirmation_token"], executor=executor
        )
    assert executor.calls == []
