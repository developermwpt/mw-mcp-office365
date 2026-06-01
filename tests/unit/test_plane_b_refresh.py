"""Unit — Plano B (T6): troca de code, refresh, invalid_grant->reauth + log, consent."""

from __future__ import annotations

import logging

import pytest

from mcp_o365.auth.errors import ConsentRequired, InvalidGrant
from mcp_o365.auth.plane_b import PlaneB
from tests.conftest import (
    CONSENT_REQUIRED_RESPONSE,
    INVALID_GRANT_RESPONSE,
    FakeMsalApp,
    graph_token_response,
)


def _plane_b(config, clock, *, code_result=None, refresh_result=None) -> PlaneB:
    fake = FakeMsalApp(code_result=code_result, refresh_result=refresh_result)
    return PlaneB(config, msal_app_factory=lambda: fake, clock=clock)


def test_build_authorization_url(config, clock):
    pb = _plane_b(config, clock)
    url = pb.build_authorization_url(state="st-1", redirect_uri="https://mcp.example.com/callback")
    assert "state=st-1" in url


def test_exchange_code_devolve_tokens(config, clock):
    pb = _plane_b(config, clock, code_result=graph_token_response())
    res = pb.exchange_code(code="abc", redirect_uri="https://mcp.example.com/callback")
    assert res.access_token == "graph-access-1"
    assert res.refresh_token == "rt-1"
    assert res.home_account_id == "user-oid-1"
    assert res.tenant_id == "tenant-test"
    # expires_at calculado a partir do relógio fixo + expires_in.
    assert res.expires_at is not None


def test_refresh_renova(config, clock):
    pb = _plane_b(config, clock, refresh_result=graph_token_response(refresh_token="rt-2"))
    res = pb.refresh(refresh_token="rt-1")
    assert res.access_token == "graph-access-1"
    assert res.refresh_token == "rt-2"


def test_invalid_grant_levanta_e_loga(config, clock, caplog):
    pb = _plane_b(config, clock, refresh_result=INVALID_GRANT_RESPONSE)
    with caplog.at_level(logging.WARNING, logger="mcp_o365.auth.plane_b"):
        with pytest.raises(InvalidGrant):
            pb.refresh(refresh_token="rt-x", subject_for_log="subj-1", account_id_for_log="acc-1")
    # Emitiu o sinal-chave para diagnosticar a Conditional Access.
    assert any(getattr(r, "fields", {}).get("event") == "refresh_failure"
               for r in caplog.records)


def test_consent_required(config, clock):
    pb = _plane_b(config, clock, code_result=CONSENT_REQUIRED_RESPONSE)
    with pytest.raises(ConsentRequired):
        pb.exchange_code(code="abc", redirect_uri="https://mcp.example.com/callback")
