"""Helper partilhado de resolução do fuso do mailbox (best-effort).

Extraído de `tools/calendar.py` (Fase 2, D1) para ser reutilizado também por `tools/email.py`
(agendamento de envio, US-1.9/1.10/1.11) sem criar um import circular entre os dois módulos
(`calendar.py` já importa `_domains` de `email.py`). A função mantém a assinatura keyword-only
e a semântica best-effort/degradação para `None` originais.
"""

from __future__ import annotations

from collections.abc import Callable
from datetime import datetime

from ..auth.errors import ReauthRequired, UpstreamAuthError
from ..auth.plane_b import PlaneB
from ..graph.client import GraphClient
from ..identity.mapping import IdentityMapping
from ..storage.token_store import TokenStore
from ._session import resolve_access_token


async def _resolve_tz(
    subject: str | None,
    *,
    mapping: IdentityMapping,
    plane_b: PlaneB,
    store: TokenStore,
    graph_client: GraphClient,
    account_id: str | None,
    clock: Callable[[], datetime],
) -> str | None:
    """D1 — lê o fuso do mailbox uma vez por pedido. None -> Graph usa UTC (sem `Prefer`).

    BEST-EFFORT por desenho: ler o fuso é acessório e exige o scope `MailboxSettings.Read`.
    Se esse scope faltar (Graph devolve 403 -> `UpstreamAuthError`) ou a sessão não resolver,
    devolve-se `None` (fallback para UTC) em vez de propagar. Nunca se passa por `call_graph`
    aqui, para que uma falha do fuso NÃO force refresh nem marque a conta como expirada —
    senão um 403 numa leitura secundária derrubava a sessão inteira (email incluído)."""
    try:
        _, token = await resolve_access_token(
            subject, mapping=mapping, plane_b=plane_b, store=store,
            account_id=account_id, clock=clock,
        )
        return await graph_client.get_mailbox_timezone(token)
    except (UpstreamAuthError, ReauthRequired):
        # Fuso indisponível (scope em falta ou sessão por reautenticar) -> UTC.
        # A reautenticação genuína, se necessária, será sinalizada pela chamada principal
        # (ex.: calendarView), que usa scopes efetivamente concedidos.
        return None
