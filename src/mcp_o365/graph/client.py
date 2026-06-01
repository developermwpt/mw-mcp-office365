"""T10 — Wrapper fino sobre o Microsoft Graph.

Na PoC só expõe `me()` (`GET /me`). Não conhece tokens nem store — recebe o access token
pronto. Trata 401/403 (erro de auth tipado) e 429 (respeita `Retry-After`). O cliente HTTP
e o `sleeper` são injetados para testes determinísticos (sem rede nem `sleep` reais).
"""

from __future__ import annotations

import asyncio
from collections.abc import Awaitable, Callable

import httpx

from ..auth.errors import UpstreamAuthError

DEFAULT_GRAPH_BASE = "https://graph.microsoft.com/v1.0"
_MAX_RETRIES_429 = 3


async def _real_sleeper(seconds: float) -> None:
    await asyncio.sleep(seconds)


class GraphError(Exception):
    """Erro genérico do Graph (não relacionado com auth)."""


class GraphClient:
    """Cliente HTTP mínimo para o Microsoft Graph."""

    def __init__(
        self,
        http: httpx.AsyncClient,
        base_url: str = DEFAULT_GRAPH_BASE,
        sleeper: Callable[[float], Awaitable[None]] = _real_sleeper,
    ) -> None:
        self._http = http
        self._base = base_url.rstrip("/")
        self._sleep = sleeper

    async def me(self, access_token: str) -> dict:
        """`GET /me` — devolve a identidade do utilizador autenticado."""
        data = await self._get("/me", access_token)
        return {
            "id": data.get("id"),
            "displayName": data.get("displayName"),
            "userPrincipalName": data.get("userPrincipalName"),
        }

    async def _get(self, path: str, access_token: str) -> dict:
        url = f"{self._base}{path}"
        headers = {"Authorization": f"Bearer {access_token}"}
        attempts = 0
        while True:
            resp = await self._http.get(url, headers=headers)
            if resp.status_code == 429 and attempts < _MAX_RETRIES_429:
                retry_after = float(resp.headers.get("Retry-After", "1"))
                await self._sleep(retry_after)
                attempts += 1
                continue
            if resp.status_code in (401, 403):
                raise UpstreamAuthError(
                    f"Graph rejeitou o token ({resp.status_code})."
                )
            if resp.status_code >= 400:
                raise GraphError(f"Graph devolveu {resp.status_code}: {resp.text[:200]}")
            return resp.json()
