"""T12 — Composition root.

Único ponto onde os componentes se instanciam e ligam: configuração, cifra, token store,
mapeamento de identidade, Plano B (Entra), provider do Plano A, cliente Graph e o servidor
MCP. Expõe `build_app()` (ASGI, para uvicorn/Cloudflare) e `main()` (arranque direto).
"""

from __future__ import annotations

import logging

import httpx

from .approval.engine import ApprovalEngine
from .auth.plane_a import MwOAuthProvider
from .auth.plane_b import PlaneB
from .config import Settings, load_settings
from .graph.client import GraphClient
from .identity.mapping import IdentityMapping
from .logging_setup import setup_logging
from .server import build_server
from .storage.crypto import LocalAesGcmCipher
from .storage.token_store import TokenStore

logger = logging.getLogger("mcp_o365.app")


def build_components(config: Settings) -> dict:
    """Instancia e liga todos os componentes a partir da configuração."""
    cipher = LocalAesGcmCipher(config.encryption_key_bytes())
    store = TokenStore(config.token_store_path, cipher)
    mapping = IdentityMapping(store)
    plane_b = PlaneB(config)
    provider = MwOAuthProvider(
        store=store, plane_b=plane_b, mapping=mapping, config=config
    )
    graph_client = GraphClient(httpx.AsyncClient(timeout=30.0))
    approval = ApprovalEngine(store, ttl_seconds=config.approval_ttl_seconds)
    server = build_server(
        config=config,
        provider=provider,
        mapping=mapping,
        plane_b=plane_b,
        graph_client=graph_client,
        store=store,
        approval=approval,
    )
    return {
        "config": config,
        "cipher": cipher,
        "store": store,
        "mapping": mapping,
        "plane_b": plane_b,
        "provider": provider,
        "graph_client": graph_client,
        "approval": approval,
        "server": server,
    }


def build_app():
    """Constrói a aplicação ASGI (streamable HTTP) pronta para servir."""
    config = load_settings()
    setup_logging(config.log_level)
    components = build_components(config)
    logger.info(
        "servidor pronto",
        extra={"fields": {"event": "startup", "host": config.bind_host, "port": config.bind_port}},
    )
    return components["server"].streamable_http_app()


def main() -> None:
    """Arranque direto via uvicorn (também há `mcp.run`, mas explicitamos o host/port)."""
    config = load_settings()
    setup_logging(config.log_level)
    components = build_components(config)
    components["server"].run(transport="streamable-http")


if __name__ == "__main__":
    main()
