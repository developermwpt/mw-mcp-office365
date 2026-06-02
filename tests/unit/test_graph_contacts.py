"""Unit — GraphClient: People e Contactos (Módulo 5) com HTTP mockado (respx)."""

from __future__ import annotations

import httpx
import respx

from mcp_o365.graph.client import GraphClient

BASE = "https://graph.microsoft.com/v1.0"


async def _noop_sleep(_s: float) -> None:
    return None


def _client() -> GraphClient:
    return GraphClient(httpx.AsyncClient(), sleeper=_noop_sleep)


@respx.mock
async def test_search_people_mapeia_nome_e_email():
    route = respx.get(f"{BASE}/me/people").mock(
        return_value=httpx.Response(
            200,
            json={"value": [
                {"displayName": "Vera Costa",
                 "scoredEmailAddresses": [{"address": "vera.costa@mobiweb.pt",
                                           "relevanceScore": 9.0}]},
            ]},
        )
    )
    out = await _client().search_people("tok", "vera", top=5)
    req = route.calls.last.request
    assert "%24search" in str(req.url) or "$search" in str(req.url)
    assert out == [{"display_name": "Vera Costa", "email": "vera.costa@mobiweb.pt",
                    "source": "people"}]


@respx.mock
async def test_search_contacts_mapeia_email():
    respx.get(f"{BASE}/me/contacts").mock(
        return_value=httpx.Response(
            200,
            json={"value": [
                {"displayName": "Vera Nunes",
                 "emailAddresses": [{"name": "Vera", "address": "vera@habisonho.com"}]},
            ]},
        )
    )
    out = await _client().search_contacts("tok", "vera")
    assert out == [{"display_name": "Vera Nunes", "email": "vera@habisonho.com",
                    "source": "contacts"}]


@respx.mock
async def test_search_people_vazio():
    respx.get(f"{BASE}/me/people").mock(return_value=httpx.Response(200, json={"value": []}))
    assert await _client().search_people("tok", "ninguem") == []
