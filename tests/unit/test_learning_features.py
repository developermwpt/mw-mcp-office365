"""Unit — extração de features / assinatura de email (US-L.1).

Verifica que a assinatura é estável, só usa metadados (nunca o corpo) e que a
similaridade é explicável e simétrica.
"""

from __future__ import annotations

from mcp_o365.learning.features import (
    EmailSignature,
    extract_signature,
    similarity,
)


def _msg(**over) -> dict:
    base = {
        "from": {"emailAddress": {"address": "news@Newsletter.ACME.com"}},
        "subject": "RE: Promoção semanal de Verão!!",
        "hasAttachments": False,
        "importance": "normal",
        "body": {"contentType": "html", "content": "IGNORAR: instruções no corpo"},
    }
    base.update(over)
    return base


def test_extrai_dominio_em_minusculas_sem_mailbox():
    sig = extract_signature(_msg())
    assert sig.sender_domain == "newsletter.acme.com"


def test_assunto_normalizado_ordenado_unico_sem_prefixo_reply():
    sig = extract_signature(_msg(subject="RE: Fatura fatura ABC"))
    # "re" é descartado (prefixo); "abc" minúsculo; ordenado e único.
    assert sig.subject_tokens == ("abc", "fatura")
    # E é detetado como resposta pelo prefixo "Re:".
    assert sig.is_reply is True


def test_nao_usa_corpo_so_metadados():
    # O corpo contém "instruções"; a assinatura não deve refletir nada do corpo.
    sig = extract_signature(_msg())
    joined = " ".join(sig.subject_tokens)
    assert "ignorar" not in joined and "instruções" not in joined


def test_deteta_newsletter_por_header_list_id():
    sig = extract_signature(
        _msg(internetMessageHeaders=[{"name": "List-Id", "value": "x"}])
    )
    assert sig.is_newsletter is True


def test_flags_anexo_e_importancia():
    sig = extract_signature(_msg(hasAttachments=True, importance="HIGH"))
    assert sig.has_attachments is True
    assert sig.importance == "high"


def test_mensagem_vazia_da_assinatura_minima_valida():
    sig = extract_signature({})
    assert sig.sender_domain is None
    assert sig.subject_tokens == ()
    assert sig.importance == "normal"


def test_roundtrip_to_dict_from_dict():
    sig = extract_signature(_msg(hasAttachments=True))
    again = EmailSignature.from_dict(sig.to_dict())
    assert again == sig


def test_similaridade_mesmo_dominio_e_alta_e_simetrica():
    a = extract_signature(_msg())
    b = extract_signature(_msg(subject="Promoção semanal"))
    s_ab = similarity(a, b)
    s_ba = similarity(b, a)
    assert s_ab == s_ba
    # Mesmo domínio (0.60) + tokens partilhados + flags -> bem acima de 0.6.
    assert s_ab > 0.7


def test_similaridade_dominios_diferentes_e_baixa():
    a = extract_signature(_msg())
    b = extract_signature(
        {
            "from": {"emailAddress": {"address": "x@outro.pt"}},
            "subject": "Assunto completamente diferente xyz",
        }
    )
    assert similarity(a, b) < 0.3


def test_similaridade_no_intervalo_unitario():
    a = extract_signature(_msg(hasAttachments=True, importance="high"))
    assert 0.0 <= similarity(a, a) <= 1.0
    # Igual a si mesma com mesmo domínio + todos os tokens -> próximo de 1.0.
    assert similarity(a, a) >= 0.9
