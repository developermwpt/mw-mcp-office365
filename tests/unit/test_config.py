"""Unit — configuração (T1): falha-rápido sem segredo, segredos fora do repr."""

from __future__ import annotations

import base64
import os

import pytest
from pydantic import ValidationError

from mcp_o365.config import Settings


def _base_env() -> dict:
    return {
        "ENTRA_TENANT_ID": "t", "ENTRA_CLIENT_ID": "c", "ENTRA_CLIENT_SECRET": "s",
        "ENTRA_AUTHORITY": "https://login.microsoftonline.com/t",
        "OAUTH_REDIRECT_URI": "https://mcp.example.com/callback",
        "GRAPH_SCOPES": "User.Read offline_access",
        "MCP_ISSUER_URL": "https://mcp.example.com",
        "MCP_PUBLIC_BASE_URL": "https://mcp.example.com",
        "TOKEN_STORE_PATH": ":memory:",
        "TOKEN_ENCRYPTION_KEY": base64.b64encode(b"\x00" * 32).decode(),
    }


def _with_env(env: dict):
    old = dict(os.environ)
    os.environ.clear()
    os.environ.update(env)
    try:
        return Settings(_env_file=None)
    finally:
        os.environ.clear()
        os.environ.update(old)


def test_config_valida_carrega():
    cfg = _with_env(_base_env())
    assert cfg.graph_scopes == ["User.Read", "offline_access"]
    assert len(cfg.encryption_key_bytes()) == 32


def test_falta_segredo_falha_rapido():
    env = _base_env()
    del env["ENTRA_CLIENT_SECRET"]
    with pytest.raises(ValidationError):
        _with_env(env)


def test_segredo_nunca_no_repr():
    cfg = _with_env(_base_env())
    assert "s" == cfg.entra_client_secret.get_secret_value()
    assert "fake" not in repr(cfg)
    # O SecretStr nunca expõe o valor em repr/str.
    assert cfg.entra_client_secret.get_secret_value() not in repr(cfg.entra_client_secret)


def test_authority_tem_de_ser_https():
    env = _base_env()
    env["ENTRA_AUTHORITY"] = "http://inseguro"
    with pytest.raises(ValidationError):
        _with_env(env)


def test_chave_de_cifra_com_tamanho_errado():
    env = _base_env()
    env["TOKEN_ENCRYPTION_KEY"] = base64.b64encode(b"curta").decode()
    cfg = _with_env(env)
    with pytest.raises(ValueError):
        cfg.encryption_key_bytes()
