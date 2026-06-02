"""Fase Aprendizagem — Tools de aprendizagem de comportamento (US-L.x).

Funções `run_*` standalone, independentes do transporte MCP, no mesmo padrão de
`tools/email.py` (dependências injetadas, `subject` do Plano A, relógio injetável).

Todas as tools são READ-ONLY no que toca ao Microsoft Graph: a aprendizagem trabalha
sobre metadados locais e sobre o store de comportamento — **não há novas chamadas Graph
nos caminhos de produção**. Executar uma recomendação é responsabilidade do utilizador e
passa SEMPRE pelo two-phase approval (`*_prepare` -> `confirmation_token` -> `*_confirm`).

Garantias (ver doc §2 e §5):
- **Opt-in obrigatório**: sem consentimento não se regista nem se recomenda; a resposta
  explica como ativar. Default: desligado.
- **Isolamento por `subject`**: todas as queries filtram por subject.
- **Auditoria só-metadados**: `learning.recommend`, `learning.event_recorded`,
  `learning.opt_in`, `learning.forget` — sem PII (subject_hash, contagens, domínio).
- **Degradação graciosa**: erros do store nunca quebram a operação principal de email.
- **Nunca auto-executar**: estas tools só sugerem.
"""

from __future__ import annotations

import logging
from collections.abc import Callable
from datetime import datetime, timezone

from ..learning.events import build_behavior_event
from ..learning.features import extract_signature
from ..learning.recommender import Recommender
from ..observability.audit import log_audit
from ..storage.token_store import TokenStore

logger = logging.getLogger("mcp_o365.tools.learning")
audit_logger = logging.getLogger("mcp_o365.audit")

_OPT_OUT_MESSAGE = (
    "A aprendizagem de comportamento está desligada (opt-out). Para ativar e passar a "
    "receber recomendações, use a ferramenta learning_opt_in com enabled=True. "
    "Só são guardados metadados (nunca o corpo dos emails) e pode apagar tudo a qualquer "
    "momento com learning_forget."
)


def _utcnow() -> datetime:
    return datetime.now(timezone.utc)


# ============================ RECOMENDAÇÕES (read-only) ============================


async def run_email_recommendations(
    subject: str | None,
    *,
    store: TokenStore,
    recommender: Recommender,
    message: dict,
    message_id: str | None = None,
    clock: Callable[[], datetime] = _utcnow,
) -> dict:
    """US-L.2 — Devolve sugestões de ação para um email (NÃO executa nada).

    `message` é o dict de metadados do email no formato Graph (tipicamente o que
    `email_read` já devolveu — sem nova chamada à rede). Cada sugestão indica o
    `prepare_tool`/`prepare_params` a usar se o utilizador aceitar: aceitar uma sugestão
    desemboca no `*_prepare` correspondente, que devolve o `confirmation_token`.
    """
    if not subject:
        # Sem sessão não há histórico nem isolamento possível — resposta amigável.
        return {"status": "ok", "recommendations": [], "count": 0,
                "note": "Sessão não autenticada."}

    if not store.get_learning_opt_in(subject):
        return {"status": "opt_out", "recommendations": [], "count": 0,
                "message": _OPT_OUT_MESSAGE}

    target = extract_signature(message)
    try:
        # Restringe ao domínio do remetente para a recomendação ser barata e focada;
        # se não houver domínio, cai no histórico geral do subject.
        events = store.list_behavior_events(subject, sender_domain=target.sender_domain)
        if not events and target.sender_domain is not None:
            events = store.list_behavior_events(subject)
    except Exception:  # noqa: BLE001 — degradação graciosa, nunca exceção crua
        logger.exception("falha ao ler histórico de comportamento")
        events = []

    recs = recommender.recommend(
        target=target, events=events, message_id=message_id
    )

    log_audit(
        audit_logger, action="learning.recommend", subject=subject,
        outcome="success",
        extra={"sender_domain": target.sender_domain, "count": len(recs),
               "history_size": len(events)},
    )
    return {
        "status": "ok",
        "recommendations": [r.to_dict() for r in recs],
        "count": len(recs),
        "note": (
            "Estas são sugestões. Para executar uma, chame o prepare_tool indicado com os "
            "prepare_params (acrescentando o conteúdo, ex.: comment numa resposta) e "
            "confirme com o *_confirm. Nada é executado automaticamente."
        ),
    }


# ============================ CONSENTIMENTO / ESQUECIMENTO ============================


async def run_learning_opt_in(
    subject: str | None,
    *,
    store: TokenStore,
    enabled: bool = True,
    clock: Callable[[], datetime] = _utcnow,
) -> dict:
    """US-L.5 — Ativa (opt-in) ou desativa (opt-out) a aprendizagem para o subject."""
    if not subject:
        return {"status": "error", "message": "Sessão não autenticada. Inicie sessão."}
    try:
        store.set_learning_opt_in(subject, enabled)
    except Exception:  # noqa: BLE001
        logger.exception("falha ao gravar opt-in de aprendizagem")
        return {"status": "error",
                "message": "Não foi possível gravar a preferência. Tente novamente."}
    log_audit(
        audit_logger, action="learning.opt_in", subject=subject, outcome="success",
        extra={"enabled": bool(enabled)},
    )
    if enabled:
        msg = (
            "Aprendizagem ativada. A partir de agora, as ações que confirmar (mover, "
            "responder, eliminar, etc.) registam metadados para gerar recomendações. "
            "Pode desativar com enabled=False e apagar o histórico com learning_forget."
        )
    else:
        msg = (
            "Aprendizagem desativada. Deixam de ser registados novos comportamentos. "
            "O histórico existente mantém-se até o apagar com learning_forget."
        )
    return {"status": "ok", "opt_in": bool(enabled), "message": msg}


async def run_learning_forget(
    subject: str | None,
    *,
    store: TokenStore,
    clock: Callable[[], datetime] = _utcnow,
) -> dict:
    """US-L.6 — Apaga TODO o histórico de comportamento do subject (esquecimento)."""
    if not subject:
        return {"status": "error", "message": "Sessão não autenticada. Inicie sessão."}
    try:
        deleted = store.purge_behavior_events(subject)
    except Exception:  # noqa: BLE001
        logger.exception("falha ao apagar histórico de comportamento")
        return {"status": "error",
                "message": "Não foi possível apagar o histórico. Tente novamente."}
    log_audit(
        audit_logger, action="learning.forget", subject=subject, outcome="success",
        extra={"deleted": deleted},
    )
    return {
        "status": "ok",
        "deleted": deleted,
        "message": f"Apagados {deleted} evento(s) de comportamento. Histórico limpo.",
    }


# ===================== REGISTO DE COMPORTAMENTO (chamado pelos confirms) =====================


def record_action_event(
    subject: str | None,
    *,
    store: TokenStore,
    action: str,
    message: dict | None = None,
    destination: str | None = None,
    clock: Callable[[], datetime] = _utcnow,
) -> None:
    """Regista, de forma DEFENSIVA, um evento de comportamento após uma ação confirmada.

    Pensada para ser chamada pelos `*_confirm` de email. Garantias:
    - **Só regista se o opt-in estiver ligado** (default desligado).
    - **Só metadados**: usa `build_behavior_event` (assinatura, domínio, destino).
    - **Nunca levanta**: qualquer falha é apenas registada em log — NÃO pode quebrar a
      operação de email principal (degradação graciosa, doc §7).
    - **Local**: não toca no Graph.
    """
    if not subject:
        return
    try:
        if not store.get_learning_opt_in(subject):
            return
        event = build_behavior_event(
            action=action, message=message, destination=destination
        )
        if event is None:
            return
        store.record_behavior_event(subject=subject, **event.as_record())
        log_audit(
            audit_logger, action="learning.event_recorded", subject=subject,
            outcome="success",
            extra={"behavior_action": action, "sender_domain": event.sender_domain},
        )
    except Exception:  # noqa: BLE001 — nunca propagar; a operação de email já concluiu
        logger.exception("falha ao registar evento de comportamento (ignorado)")
