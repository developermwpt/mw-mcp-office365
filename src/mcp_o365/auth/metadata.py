"""T7 — Metadata OAuth do Plano A.

# NOTA SDK: o `mcp` 1.27.2 serve nativamente os documentos `.well-known` exigidos
# (RFC 9728 protected-resource + RFC 8414 authorization-server) quando se passa um
# `AuthSettings` + `auth_server_provider` ao `FastMCP`. Por isso NÃO reimplementamos os
# `.well-known` à mão (como o plano original assumia) — este módulo é o adaptador fino
# que deriva o `AuthSettings` da configuração. Cumpre a divergência aprovada.
"""

from __future__ import annotations

from mcp.server.auth.settings import AuthSettings
from pydantic import AnyHttpUrl

from ..config import Settings
from .dcr import build_client_registration_options


def build_auth_settings(config: Settings) -> AuthSettings:
    """Constrói o `AuthSettings` que faz o SDK montar a metadata e as rotas OAuth."""
    return AuthSettings(
        issuer_url=AnyHttpUrl(config.mcp_issuer_url),
        resource_server_url=AnyHttpUrl(config.mcp_public_base_url),
        required_scopes=[],
        client_registration_options=build_client_registration_options(config),
    )
