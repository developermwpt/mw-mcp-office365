"""T11/T12 — Definição do servidor MCP (FastMCP) e ligação das rotas.

Compõe o `FastMCP` com o `AuthSettings` (Plano A via SDK), regista a tool `whoami` e monta
as rotas custom `/callback` (perna do Entra — Plano B) e `/healthz`/`/readyz`.

# NOTA SDK: passamos apenas `auth_server_provider` (não `token_verifier`) — o SDK rejeita
# os dois em simultâneo e deriva o verifier do provider automaticamente (ProviderTokenVerifier).
"""

from __future__ import annotations

import logging

from mcp.server.auth.middleware.auth_context import get_access_token
from mcp.server.fastmcp import FastMCP
from starlette.requests import Request
from starlette.responses import JSONResponse, RedirectResponse

from .auth.errors import AuthError
from .auth.metadata import build_auth_settings
from .auth.plane_a import MwOAuthProvider
from .auth.plane_b import PlaneB
from .config import Settings
from .graph.client import GraphClient
from .identity.mapping import IdentityMapping
from .observability import health
from .storage.token_store import TokenStore
from .tools.whoami import run_whoami

logger = logging.getLogger("mcp_o365.server")


def build_server(
    *,
    config: Settings,
    provider: MwOAuthProvider,
    mapping: IdentityMapping,
    plane_b: PlaneB,
    graph_client: GraphClient,
    store: TokenStore,
) -> FastMCP:
    """Constrói e devolve o servidor MCP totalmente ligado."""
    mcp = FastMCP(
        name="mw-mcp-office365",
        instructions=(
            "Servidor MCP de Office 365 (PoC Fase 0). Disponível: a ferramenta read-only "
            "`whoami`, que confirma a identidade do utilizador autenticado via Microsoft Graph."
        ),
        auth=build_auth_settings(config),
        auth_server_provider=provider,
        host=config.bind_host,
        port=config.bind_port,
    )

    @mcp.tool(description="Devolve a identidade do utilizador O365 autenticado (read-only).")
    async def whoami() -> dict:
        token = get_access_token()
        subject = token.subject if token else None
        return await run_whoami(
            subject,
            mapping=mapping,
            plane_b=plane_b,
            graph_client=graph_client,
            store=store,
        )

    @mcp.custom_route("/callback", methods=["GET"], name="entra_callback")
    async def entra_callback(request: Request) -> RedirectResponse | JSONResponse:
        """Callback do Entra (Plano B): conclui o login e volta ao Claude."""
        error = request.query_params.get("error")
        if error:
            desc = request.query_params.get("error_description", error)
            return JSONResponse({"error": error, "error_description": desc}, status_code=400)
        code = request.query_params.get("code")
        state = request.query_params.get("state")
        if not code or not state:
            return JSONResponse({"error": "missing code/state"}, status_code=400)
        try:
            redirect_url = provider.complete_entra_callback(code=code, state=state)
        except AuthError as exc:
            return JSONResponse({"error": "auth_error", "detail": str(exc)}, status_code=400)
        return RedirectResponse(url=redirect_url, status_code=302)

    @mcp.custom_route("/healthz", methods=["GET"], name="healthz")
    async def healthz(_request: Request) -> JSONResponse:
        return JSONResponse(health.liveness())

    @mcp.custom_route("/readyz", methods=["GET"], name="readyz")
    async def readyz(_request: Request) -> JSONResponse:
        ready, detail = health.readiness(store, config_loaded=True)
        return JSONResponse(detail, status_code=200 if ready else 503)

    return mcp
