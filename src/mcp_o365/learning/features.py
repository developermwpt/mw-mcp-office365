"""Fase Aprendizagem — extração de features / assinatura de um email (US-L.1).

Deriva uma "assinatura" estável e explicável a partir dos **metadados** de uma mensagem
no formato Microsoft Graph (o mesmo dict devolvido por `email_read`/`email_search`).
Funções PURAS e determinísticas, sem efeitos colaterais nem rede — testáveis isoladamente.

Privacidade / anti prompt injection (ver doc §2 e §4):
- Nunca se lê o CORPO da mensagem; só campos de metadados (remetente, assunto, flags).
- O assunto NÃO é guardado em claro: é normalizado e reduzido a um conjunto de *tokens*
  (palavras minúsculas, sem pontuação, sem stopwords curtas), suficiente para medir
  similaridade mas não para reconstruir o texto. As features são metadados, não instruções:
  o conteúdo nunca é tratado como ordem.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

# Tokens de assunto: só letras/dígitos, em minúsculas. Descartam-se tokens com < 3 chars
# (ruído: "re", "fw", "de"...) para a assinatura ser estável entre variações triviais.
_TOKEN_RE = re.compile(r"[a-z0-9]+")
_MIN_TOKEN_LEN = 3
# Prefixos de assunto de resposta/reencaminho a ignorar (vários idiomas).
_REPLY_PREFIXES = ("re", "re:", "fw", "fwd", "enc", "res")


def _normalize_subject_tokens(subject: str | None) -> tuple[str, ...]:
    """Reduz o assunto a um conjunto ordenado de tokens normalizados e únicos."""
    if not subject:
        return ()
    tokens = [
        t
        for t in _TOKEN_RE.findall(subject.lower())
        if len(t) >= _MIN_TOKEN_LEN and t not in _REPLY_PREFIXES
    ]
    # Ordenado + único: a assinatura não depende da ordem das palavras.
    return tuple(sorted(set(tokens)))


def _sender_address(message: dict) -> str | None:
    """Extrai o endereço do remetente (`from`/`sender`) em minúsculas, se existir."""
    for key in ("from", "sender"):
        node = message.get(key) or {}
        addr = (node.get("emailAddress") or {}).get("address")
        if addr:
            return addr.strip().lower()
    return None


def _domain_of(address: str | None) -> str | None:
    """Domínio de um endereço (parte após o @), sem a mailbox local (menos PII)."""
    if address and "@" in address:
        return address.rsplit("@", 1)[1].lower()
    return None


@dataclass(frozen=True)
class EmailSignature:
    """Assinatura imutável de um email — só metadados, sem corpo nem PII em claro.

    `sender_domain` é o sinal mais forte e estável; os `subject_tokens` afinam a
    similaridade dentro do mesmo domínio. As flags são metadados booleanos do Graph.
    """

    sender_domain: str | None
    subject_tokens: tuple[str, ...] = field(default_factory=tuple)
    has_attachments: bool = False
    is_reply: bool = False
    importance: str = "normal"  # low | normal | high
    is_newsletter: bool = False  # presença de cabeçalho List-Id / List-Unsubscribe

    def to_dict(self) -> dict:
        """Serializa para JSON (guardado cifrado no store; só metadados)."""
        return {
            "sender_domain": self.sender_domain,
            "subject_tokens": list(self.subject_tokens),
            "has_attachments": self.has_attachments,
            "is_reply": self.is_reply,
            "importance": self.importance,
            "is_newsletter": self.is_newsletter,
        }

    @classmethod
    def from_dict(cls, data: dict) -> EmailSignature:
        """Reconstrói a assinatura a partir do dict guardado no store."""
        return cls(
            sender_domain=data.get("sender_domain"),
            subject_tokens=tuple(data.get("subject_tokens") or ()),
            has_attachments=bool(data.get("has_attachments")),
            is_reply=bool(data.get("is_reply")),
            importance=str(data.get("importance") or "normal"),
            is_newsletter=bool(data.get("is_newsletter")),
        )


def _detect_newsletter(message: dict) -> bool:
    """Deteta newsletter/lista por cabeçalhos `List-*` (metadados, não conteúdo).

    Aceita tanto `internetMessageHeaders` (lista do Graph) como uma flag já derivada
    `is_newsletter`/`list_id` para facilitar a construção em testes.
    """
    if message.get("is_newsletter") or message.get("list_id"):
        return True
    headers = message.get("internetMessageHeaders") or []
    for h in headers:
        name = (h.get("name") or "").lower()
        if name in ("list-id", "list-unsubscribe"):
            return True
    return False


def extract_signature(message: dict) -> EmailSignature:
    """Deriva a `EmailSignature` a partir de um dict de mensagem no formato Graph.

    Só toca em campos de METADADOS — nunca em `body`. É a fronteira anti prompt
    injection: o que aqui entra é tratado como dado, nunca como instrução.
    """
    address = _sender_address(message)
    has_attachments = bool(message.get("hasAttachments"))
    is_reply = bool(
        message.get("inReplyTo")
        or message.get("conversationId") and message.get("isReply")
    )
    # Heurística adicional de resposta: prefixo "Re:" no assunto bruto.
    raw_subject = message.get("subject") or ""
    if raw_subject.strip().lower().startswith(("re:", "re ", "res:")):
        is_reply = True
    importance = str(message.get("importance") or "normal").lower()
    if importance not in ("low", "normal", "high"):
        importance = "normal"
    return EmailSignature(
        sender_domain=_domain_of(address),
        subject_tokens=_normalize_subject_tokens(raw_subject),
        has_attachments=has_attachments,
        is_reply=is_reply,
        importance=importance,
        is_newsletter=_detect_newsletter(message),
    )


def similarity(a: EmailSignature, b: EmailSignature) -> float:
    """Score de similaridade explicável entre duas assinaturas, em [0, 1].

    Combinação linear de sinais simples e auditáveis (ver doc §4):
    - domínio do remetente igual: peso 0.60 (sinal dominante e estável);
    - sobreposição de tokens de assunto (Jaccard): peso 0.25;
    - mesma flag de anexo: peso 0.05; mesma flag de resposta: peso 0.05;
    - mesma flag de newsletter: peso 0.05.

    A fórmula é deliberadamente linear e transparente: cada parcela é inspecionável e
    o `rationale` da recomendação consegue explicá-la em linguagem natural.
    """
    score = 0.0
    if a.sender_domain and a.sender_domain == b.sender_domain:
        score += 0.60
    score += 0.25 * _jaccard(set(a.subject_tokens), set(b.subject_tokens))
    if a.has_attachments == b.has_attachments:
        score += 0.05
    if a.is_reply == b.is_reply:
        score += 0.05
    if a.is_newsletter == b.is_newsletter:
        score += 0.05
    return round(min(score, 1.0), 4)


def _jaccard(a: set[str], b: set[str]) -> float:
    """Índice de Jaccard entre dois conjuntos; 0.0 se ambos vazios."""
    if not a and not b:
        return 0.0
    union = a | b
    if not union:
        return 0.0
    return len(a & b) / len(union)
