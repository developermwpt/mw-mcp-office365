"""Unit — token store (T4): cifra em repouso, isolamento por subject, multi-conta, expiração."""

from __future__ import annotations

from datetime import timedelta

from tests.conftest import FIXED_NOW


def test_token_cifrado_em_repouso(store):
    sid = store.create_session("subject-A")
    store.upsert_account(
        subject="subject-A", session_id=sid,
        access_token="AT-SECRETO", refresh_token="RT-SECRETO",
        expires_at=FIXED_NOW + timedelta(hours=1), home_account_id="acc-1",
    )
    # Lê o BLOB em bruto da coluna: não pode conter o token em claro.
    raw = store._conn.execute(
        "SELECT access_token_enc, refresh_token_enc FROM linked_accounts"
    ).fetchone()
    assert b"AT-SECRETO" not in bytes(raw[0])
    assert b"RT-SECRETO" not in bytes(raw[1])
    # Mas a leitura pela API decifra corretamente.
    acc = store.get_account("subject-A", "acc-1")
    assert acc["access_token"] == "AT-SECRETO"
    assert acc["refresh_token"] == "RT-SECRETO"


def test_isolamento_entre_utilizadores(store):
    sa = store.create_session("subject-A")
    sb = store.create_session("subject-B")
    store.upsert_account(subject="subject-A", session_id=sa, access_token="a",
                         refresh_token=None, expires_at=None, home_account_id="acc-A")
    store.upsert_account(subject="subject-B", session_id=sb, access_token="b",
                         refresh_token=None, expires_at=None, home_account_id="acc-B")
    # B não vê a conta de A, mesmo sabendo o account_id.
    assert store.get_account("subject-B", "acc-A") is None
    assert store.get_account("subject-A", "acc-A") is not None
    assert {a["account_id"] for a in store.list_accounts("subject-A")} == {"acc-A"}


def test_multi_conta_na_mesma_sessao(store):
    sid = store.create_session("subject-A")
    store.upsert_account(subject="subject-A", session_id=sid, access_token="a1",
                         refresh_token=None, expires_at=None, home_account_id="acc-1",
                         is_default=True)
    store.upsert_account(subject="subject-A", session_id=sid, access_token="a2",
                         refresh_token=None, expires_at=None, home_account_id="acc-2")
    accounts = store.list_accounts("subject-A")
    assert len(accounts) == 2
    assert store.get_default_account("subject-A")["account_id"] == "acc-1"


def test_default_unico_ao_marcar_novo(store):
    sid = store.create_session("subject-A")
    store.upsert_account(subject="subject-A", session_id=sid, access_token="a1",
                         refresh_token=None, expires_at=None, home_account_id="acc-1",
                         is_default=True)
    store.upsert_account(subject="subject-A", session_id=sid, access_token="a2",
                         refresh_token=None, expires_at=None, home_account_id="acc-2",
                         is_default=True)
    defaults = [a for a in store.list_accounts("subject-A") if a["is_default"]]
    assert len(defaults) == 1
    assert defaults[0]["account_id"] == "acc-2"


def test_marcar_expirada_remove_de_ativas(store):
    sid = store.create_session("subject-A")
    store.upsert_account(subject="subject-A", session_id=sid, access_token="a",
                         refresh_token=None, expires_at=None, home_account_id="acc-1")
    store.mark_session_expired("subject-A")
    assert store.get_active_session("subject-A") is None
    assert store.list_accounts("subject-A") == []


def test_update_tokens_so_afeta_conta_certa(store):
    sid = store.create_session("subject-A")
    store.upsert_account(subject="subject-A", session_id=sid, access_token="old",
                         refresh_token="rt", expires_at=None, home_account_id="acc-1")
    store.update_account_tokens(subject="subject-A", account_id="acc-1",
                                access_token="new", refresh_token=None,
                                expires_at=FIXED_NOW + timedelta(hours=1))
    acc = store.get_account("subject-A", "acc-1")
    assert acc["access_token"] == "new"
    assert acc["refresh_token"] == "rt"  # COALESCE preserva o refresh quando None
