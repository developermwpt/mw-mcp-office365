"""T12 — Health checks.

`liveness` é trivial (o processo responde). `readiness` confirma que a DB está acessível e
a configuração carregada. Nenhum dos dois toca no Entra/Graph — não consomem rate limit
nem dependem de serviços externos (v1.1 §8).
"""

from __future__ import annotations

from ..storage.token_store import TokenStore


def liveness() -> dict:
    return {"status": "ok"}


def readiness(store: TokenStore, config_loaded: bool) -> tuple[bool, dict]:
    """Devolve (ready, detalhe). `ready=False` se a DB estiver inacessível."""
    db_ok = False
    try:
        db_ok = store.ping()
    except Exception:  # noqa: BLE001 — readiness não deve propagar exceção
        db_ok = False
    ready = db_ok and config_loaded
    return ready, {
        "status": "ready" if ready else "not_ready",
        "db": "ok" if db_ok else "fail",
        "config": "ok" if config_loaded else "fail",
    }
