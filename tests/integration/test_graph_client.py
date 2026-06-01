"""Integração — GraphClient (T10): me(), 401 tipado, 429 respeita Retry-After sem dormir."""

from __future__ import annotations

import httpx
import pytest
import respx

from mcp_o365.auth.errors import UpstreamAuthError
from mcp_o365.graph.client import GraphClient

GRAPH_ME = "https://graph.microsoft.com/v1.0/me"
ME_BODY = {"id": "1", "displayName": "X", "userPrincipalName": "x@e.com"}


@respx.mock
async def test_me_ok():
    respx.get(GRAPH_ME).mock(return_value=httpx.Response(200, json=ME_BODY))
    gc = GraphClient(httpx.AsyncClient())
    assert (await gc.me("tok"))["userPrincipalName"] == "x@e.com"


@respx.mock
async def test_401_tipado():
    respx.get(GRAPH_ME).mock(return_value=httpx.Response(401, json={"error": "x"}))
    gc = GraphClient(httpx.AsyncClient())
    with pytest.raises(UpstreamAuthError):
        await gc.me("tok")


@respx.mock
async def test_429_respeita_retry_after_sem_dormir():
    slept: list[float] = []

    async def fake_sleep(s: float) -> None:
        slept.append(s)

    route = respx.get(GRAPH_ME)
    route.side_effect = [
        httpx.Response(429, headers={"Retry-After": "2"}),
        httpx.Response(200, json=ME_BODY),
    ]
    gc = GraphClient(httpx.AsyncClient(), sleeper=fake_sleep)
    result = await gc.me("tok")
    assert result["id"] == "1"
    assert slept == [2.0]  # respeitou o Retry-After, mas via sleeper injetável (sem dormir)
