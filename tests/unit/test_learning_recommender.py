"""Unit — gerador de recomendações (US-L.2).

Verifica: sem histórico -> sem recomendações; com histórico consistente -> recomenda a
ação esperada com score e ordenação corretos; explicabilidade (rationale presente) e o
encadeamento no two-phase approval (prepare_tool/prepare_params).
"""

from __future__ import annotations

from mcp_o365.learning.features import EmailSignature
from mcp_o365.learning.recommender import Recommender


def _sig(domain: str = "newsletter.acme.com", tokens=("promo", "semanal")) -> EmailSignature:
    return EmailSignature(sender_domain=domain, subject_tokens=tuple(tokens))


def _event(action: str, *, destination=None, sig: EmailSignature | None = None) -> dict:
    return {
        "action": action,
        "destination": destination,
        "features": (sig or _sig()).to_dict(),
    }


def test_sem_historico_sem_recomendacoes():
    rec = Recommender()
    out = rec.recommend(target=_sig(), events=[])
    assert out == []


def test_suporte_insuficiente_nao_recomenda():
    # Apenas 1 evento (< _MIN_SUPPORT=2) -> sem recomendação.
    rec = Recommender()
    out = rec.recommend(target=_sig(), events=[_event("move", destination="Archive")])
    assert out == []


def test_recomenda_mover_para_archive_com_rationale_e_prepare():
    rec = Recommender(min_confidence=0.4)
    events = [_event("archive", destination="Archive") for _ in range(8)]
    out = rec.recommend(target=_sig(), events=events, message_id="m1")
    assert len(out) == 1
    r = out[0]
    assert r.action == "archive"
    assert r.params["destination"] == "Archive"
    assert 0.0 < r.confidence <= 1.0
    # Explicabilidade: rationale legível com domínio e contagem.
    assert "newsletter.acme.com" in r.rationale
    assert "8" in r.rationale
    # Encadeamento no two-phase approval.
    assert r.prepare_tool == "email_move_prepare"
    assert r.prepare_params == {"message_id": "m1", "destination": "Archive"}


def test_ordenacao_por_confianca_desc_e_top_n():
    # Padrão forte (8x archive) e padrão fraco (2x delete) -> archive primeiro.
    rec = Recommender(min_confidence=0.3, top_n=1)
    events = [_event("archive", destination="Archive") for _ in range(8)]
    events += [_event("delete") for _ in range(2)]
    out = rec.recommend(target=_sig(), events=events)
    assert len(out) == 1  # top_n=1
    assert out[0].action == "archive"


def test_limiar_de_confianca_filtra():
    # Mesmo com suporte, se a similaridade for baixa (domínio diferente) -> filtrado.
    rec = Recommender(min_confidence=0.6)
    target = _sig(domain="banco.pt", tokens=("extrato",))
    events = [
        _event("delete", sig=_sig(domain="spam.xyz", tokens=("promo",)))
        for _ in range(5)
    ]
    out = rec.recommend(target=target, events=events)
    assert out == []


def test_determinismo_mesma_entrada_mesma_saida():
    rec = Recommender(min_confidence=0.3)
    events = [_event("reply") for _ in range(4)]
    first = rec.recommend(target=_sig(), events=events, message_id="m1")
    second = rec.recommend(target=_sig(), events=events, message_id="m1")
    assert [r.to_dict() for r in first] == [r.to_dict() for r in second]


def test_acao_desconhecida_ignorada():
    rec = Recommender(min_confidence=0.1)
    events = [_event("flag_inexistente") for _ in range(5)]
    assert rec.recommend(target=_sig(), events=events) == []


def test_supressao_filtra_recomendacao():
    rec = Recommender(min_confidence=0.3)
    events = [_event("archive", destination="Archive") for _ in range(5)]
    sig = _sig()  # domínio newsletter.acme.com
    assert len(rec.recommend(target=sig, events=events)) == 1  # sem supressão
    out = rec.recommend(
        target=sig, events=events,
        suppressions={("newsletter.acme.com", "archive")},
    )
    assert out == []  # supressão explícita filtra


def test_recencia_reduz_confianca_de_habitos_antigos():
    from datetime import datetime, timedelta, timezone
    now = datetime(2026, 6, 1, 12, 0, 0, tzinfo=timezone.utc)
    rec = Recommender(min_confidence=0.0, half_life_days=30.0, clock=lambda: now)
    sig = _sig()

    def ev(created):
        e = _event("reply", sig=sig)
        e["created_at"] = created
        return e

    recent = [ev(now) for _ in range(4)]
    old = [ev(now - timedelta(days=120)) for _ in range(4)]
    c_recent = rec.recommend(target=sig, events=recent)[0].confidence
    c_old = rec.recommend(target=sig, events=old)[0].confidence
    assert c_recent > c_old
    # 120 dias com meia-vida 30 => ~0.5^4 = 0.0625 do peso -> bem mais baixo.
    assert c_old < c_recent * 0.2
