"""Fase Aprendizagem — normalização de eventos de comportamento (US-L.1).

Helper puro que, a partir de uma ação confirmada e dos metadados disponíveis da mensagem
(o mesmo dict do Graph já em memória — sem nova chamada à rede), constrói o registo de
comportamento a persistir: a `action`, o `sender_domain`, o `destination` (se aplicável)
e a assinatura de features (só metadados).

Não toca no store nem no Graph — apenas transforma dados já disponíveis.
"""

from __future__ import annotations

from dataclasses import dataclass

from .features import EmailSignature, extract_signature

# Ações de email cujo comportamento é aprendível.
_LEARNABLE_ACTIONS = {"move", "archive", "reply", "reply_all", "forward", "delete", "send"}


@dataclass(frozen=True)
class BehaviorEvent:
    """Evento de comportamento normalizado, pronto para `store.record_behavior_event`."""

    action: str
    signature: EmailSignature
    sender_domain: str | None
    destination: str | None

    def as_record(self) -> dict:
        """Argumentos (kwargs) para `TokenStore.record_behavior_event`."""
        return {
            "action": self.action,
            "features": self.signature.to_dict(),
            "sender_domain": self.sender_domain,
            "destination": self.destination,
        }


def build_behavior_event(
    *,
    action: str,
    message: dict | None,
    destination: str | None = None,
) -> BehaviorEvent | None:
    """Constrói um `BehaviorEvent` a partir de uma ação confirmada e dos metadados.

    Devolve `None` se a ação não for aprendível (degradação graciosa: o chamador
    simplesmente não regista nada). `message` pode ser um dict parcial do Graph (basta
    `from`/`subject`/flags) ou `None` — neste caso a assinatura fica mínima mas válida.
    """
    if action not in _LEARNABLE_ACTIONS:
        return None
    signature = extract_signature(message or {})
    return BehaviorEvent(
        action=action,
        signature=signature,
        sender_domain=signature.sender_domain,
        destination=destination,
    )
