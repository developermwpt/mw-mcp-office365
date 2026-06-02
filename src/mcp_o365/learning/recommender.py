"""Fase Aprendizagem — gerador de recomendações (US-L.2).

Dado a assinatura de um email-alvo e o histórico de comportamento do MESMO `subject`,
produz recomendações ordenadas por confiança. Cada recomendação é **explicável**
(`rationale` em linguagem natural) e **determinística** (mesma entrada -> mesma saída).

NUNCA executa nada: limita-se a sugerir uma `action` + `params`. A execução real passa
SEMPRE pelo fluxo de aprovação em duas fases existente (a tool de recomendação devolve
sugestões; aceitar uma sugestão chama o `*_prepare` correspondente).

Modelo (ver doc §4 e §5):
- Agrupam-se os eventos passados por (ação, destino) e soma-se a similaridade da assinatura
  do alvo com cada evento desse grupo. O score do grupo é a média ponderada da similaridade,
  multiplicada por um fator de suporte (quantos eventos sustentam o padrão), saturado em 1.0.
- Só entram grupos cujo score >= `min_confidence`. Devolve no máximo `top_n`.
"""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass
from datetime import datetime, timezone

from .features import EmailSignature, similarity

# Ações cujo padrão é recomendável e o `*_prepare` correspondente que as concretiza.
# `archive` é um caso especial de `move` para a pasta de arquivo.
_ACTION_TO_PREPARE = {
    "move": "email_move_prepare",
    "archive": "email_move_prepare",
    "reply": "email_reply_prepare",
    "reply_all": "email_reply_prepare",
    "forward": "email_reply_prepare",
    "delete": "email_delete_prepare",
}

# Suporte mínimo (nº de eventos) para um padrão ser considerado, e ponto de saturação:
# a partir de `_SUPPORT_SATURATION` eventos o fator de suporte é 1.0 (não cresce mais).
_MIN_SUPPORT = 2
_SUPPORT_SATURATION = 5.0


def _utcnow() -> datetime:
    return datetime.now(timezone.utc)


@dataclass(frozen=True)
class Recommendation:
    """Uma sugestão de ação, pronta a ser apresentada ao utilizador.

    `prepare_tool`/`prepare_params` indicam ao chamador qual `*_prepare` invocar se o
    utilizador aceitar — encadeando no two-phase approval, sem segundo mecanismo.
    """

    action: str
    params: dict
    confidence: float
    rationale: str
    prepare_tool: str
    prepare_params: dict

    def to_dict(self) -> dict:
        return {
            "action": self.action,
            "params": self.params,
            "confidence": self.confidence,
            "rationale": self.rationale,
            "prepare_tool": self.prepare_tool,
            "prepare_params": self.prepare_params,
        }


class Recommender:
    """Produz recomendações determinísticas e explicáveis a partir do histórico.

    Sem estado partilhado: recebe a assinatura-alvo e a lista de eventos (cada um com
    `action`, `destination`, `features`) e devolve as melhores sugestões. O relógio é
    injetável por coerência com o resto do projeto (não é usado para o score, mas fica
    disponível para futuras heurísticas de recência).
    """

    def __init__(
        self,
        *,
        min_confidence: float = 0.5,
        top_n: int = 3,
        clock: Callable[[], datetime] = _utcnow,
    ) -> None:
        self._min_confidence = min_confidence
        self._top_n = top_n
        self._clock = clock

    def recommend(
        self,
        *,
        target: EmailSignature,
        events: list[dict],
        message_id: str | None = None,
    ) -> list[Recommendation]:
        """Gera recomendações para `target` a partir do histórico `events`.

        Sem histórico (ou sem padrões acima do limiar) -> lista vazia (degradação
        graciosa: o chamador devolve uma mensagem amigável, nunca um erro).
        """
        # Agrupa por (ação, destino) acumulando a similaridade e o suporte.
        groups: dict[tuple[str, str | None], list[float]] = {}
        for ev in events:
            action = ev.get("action")
            if action not in _ACTION_TO_PREPARE:
                continue
            sig = EmailSignature.from_dict(ev.get("features") or {})
            sim = similarity(target, sig)
            if sim <= 0.0:
                continue
            key = (action, ev.get("destination"))
            groups.setdefault(key, []).append(sim)

        recs: list[Recommendation] = []
        for (action, destination), sims in groups.items():
            support = len(sims)
            if support < _MIN_SUPPORT:
                continue
            avg_sim = sum(sims) / support
            support_factor = min(support / _SUPPORT_SATURATION, 1.0)
            confidence = round(avg_sim * support_factor, 4)
            if confidence < self._min_confidence:
                continue
            recs.append(
                self._build(
                    action=action,
                    destination=destination,
                    confidence=confidence,
                    support=support,
                    target=target,
                    message_id=message_id,
                )
            )

        # Ordena por confiança desc; desempate determinístico por ação/destino.
        recs.sort(key=lambda r: (-r.confidence, r.action, str(r.params)))
        return recs[: self._top_n]

    def _build(
        self,
        *,
        action: str,
        destination: str | None,
        confidence: float,
        support: int,
        target: EmailSignature,
        message_id: str | None,
    ) -> Recommendation:
        """Monta a recomendação com `rationale` legível e os params do `*_prepare`."""
        prepare_tool = _ACTION_TO_PREPARE[action]
        domain = target.sender_domain or "este remetente"
        params: dict = {}
        prepare_params: dict = {}
        if message_id:
            prepare_params["message_id"] = message_id

        if action in ("move", "archive"):
            dest = destination or "Archive"
            params = {"destination": dest}
            prepare_params["destination"] = dest
            rationale = (
                f"Costumas mover emails de @{domain} para '{dest}' "
                f"({support} vez(es) no teu histórico)."
            )
        elif action in ("reply", "reply_all", "forward"):
            params = {"mode": action}
            prepare_params["mode"] = action
            label = {
                "reply": "responder",
                "reply_all": "responder a todos",
                "forward": "reencaminhar",
            }[action]
            rationale = (
                f"Costumas {label} a emails de @{domain} "
                f"({support} vez(es) no teu histórico)."
            )
        else:  # delete
            params = {}
            rationale = (
                f"Costumas eliminar emails de @{domain} "
                f"({support} vez(es) no teu histórico)."
            )

        return Recommendation(
            action=action,
            params=params,
            confidence=confidence,
            rationale=rationale,
            prepare_tool=prepare_tool,
            prepare_params=prepare_params,
        )
