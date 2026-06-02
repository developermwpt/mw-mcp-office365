"""Unit — store de aprendizagem (US-L.1/L.5/L.6).

Verifica: opt-in default desligado e toggle; record/list/purge isolados por subject;
features cifradas no disco; purga por data (retenção) e total (esquecimento).
"""

from __future__ import annotations


def _features(domain: str = "acme.com") -> dict:
    return {"sender_domain": domain, "subject_tokens": ["fatura"], "is_reply": False}


def test_opt_in_default_desligado_e_toggle(store):
    assert store.get_learning_opt_in("subj-1") is False
    store.set_learning_opt_in("subj-1", True)
    assert store.get_learning_opt_in("subj-1") is True
    store.set_learning_opt_in("subj-1", False)
    assert store.get_learning_opt_in("subj-1") is False


def test_record_e_list_isolados_por_subject(store):
    store.record_behavior_event(
        subject="subj-A", action="archive", features=_features(),
        sender_domain="acme.com", destination="Archive",
    )
    store.record_behavior_event(
        subject="subj-B", action="delete", features=_features("outro.pt"),
        sender_domain="outro.pt",
    )
    a = store.list_behavior_events("subj-A")
    b = store.list_behavior_events("subj-B")
    assert len(a) == 1 and a[0]["action"] == "archive"
    assert a[0]["features"]["sender_domain"] == "acme.com"
    assert len(b) == 1 and b[0]["action"] == "delete"
    # subj-A nunca vê os eventos de subj-B.
    assert all(e["sender_domain"] == "acme.com" for e in a)


def test_list_filtra_por_dominio(store):
    store.record_behavior_event(
        subject="s", action="archive", features=_features("acme.com"),
        sender_domain="acme.com", destination="Archive",
    )
    store.record_behavior_event(
        subject="s", action="delete", features=_features("spam.xyz"),
        sender_domain="spam.xyz",
    )
    only_acme = store.list_behavior_events("s", sender_domain="acme.com")
    assert len(only_acme) == 1
    assert only_acme[0]["sender_domain"] == "acme.com"


def test_features_cifradas_no_disco(store, cipher):
    store.record_behavior_event(
        subject="s", action="archive", features={"subject_tokens": ["segredo"]},
        sender_domain="acme.com",
    )
    # Lê o BLOB cru: não pode conter o token em claro.
    with store._lock:  # noqa: SLF001 — inspeção de teste
        row = store._conn.execute(
            "SELECT features_enc FROM behavior_events WHERE subject='s'"
        ).fetchone()
    assert b"segredo" not in bytes(row["features_enc"])
    # E decifra corretamente via a API pública.
    assert store.list_behavior_events("s")[0]["features"]["subject_tokens"] == ["segredo"]


def test_purge_total_apaga_so_o_subject(store):
    store.record_behavior_event(subject="a", action="delete", features=_features())
    store.record_behavior_event(subject="b", action="delete", features=_features())
    deleted = store.purge_behavior_events("a")
    assert deleted == 1
    assert store.list_behavior_events("a") == []
    assert len(store.list_behavior_events("b")) == 1


def test_purge_por_data_respeita_retencao(store, clock):
    # Evento "antigo" em T0; avança para T1; novo evento em T1. Purga before=T1 apaga
    # só o antigo (created_at < T1), preservando o recente.
    store.record_behavior_event(subject="s", action="delete", features=_features())
    clock.advance(60)
    cutoff = clock()
    store.record_behavior_event(subject="s", action="archive", features=_features())
    deleted = store.purge_behavior_events("s", before=cutoff)
    assert deleted == 1
    restantes = store.list_behavior_events("s")
    assert len(restantes) == 1
    assert restantes[0]["action"] == "archive"


def test_supressoes_isoladas_por_subject(store):
    store.add_learning_suppression("a", sender_domain="x.com", action="archive")
    store.add_learning_suppression("b", sender_domain="y.com", action="delete")
    a = store.list_learning_suppressions("a")
    assert a == [{"sender_domain": "x.com", "action": "archive"}]
    b = store.list_learning_suppressions("b")
    assert len(b) == 1 and b[0]["action"] == "delete"
    # Reinserir o mesmo par não duplica (INSERT OR REPLACE).
    store.add_learning_suppression("a", sender_domain="x.com", action="archive")
    assert len(store.list_learning_suppressions("a")) == 1
