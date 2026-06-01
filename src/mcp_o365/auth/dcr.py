"""T8 — Dynamic Client Registration (RFC 7591) do Plano A.

# NOTA SDK: o `mcp` 1.27.2 expõe o endpoint `POST /register` automaticamente quando
# `ClientRegistrationOptions(enabled=True)` é passado no `AuthSettings`, e encaminha o
# registo para `register_client()`/`get_client()` do provider. Este módulo é o adaptador
# que configura as opções de registo; a persistência real em `oauth_clients` acontece no
# provider (plane_a.py). Cumpre a divergência aprovada.
"""

from __future__ import annotations

from mcp.server.auth.settings import ClientRegistrationOptions

from ..config import Settings


def build_client_registration_options(config: Settings) -> ClientRegistrationOptions:
    """Opções de DCR para os clientes Claude remotos."""
    scopes = config.graph_scopes
    return ClientRegistrationOptions(
        enabled=True,
        valid_scopes=scopes,
        default_scopes=scopes,
    )
