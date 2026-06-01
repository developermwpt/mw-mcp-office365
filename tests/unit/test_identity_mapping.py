"""Unit — mapeamento de identidade (T5)."""

from __future__ import annotations

from datetime import timedelta

from tests.conftest import FIXED_NOW


def test_subject_desconhecido_sem_sessao(mapping):
    assert mapping.get_session("ninguem") is None
    assert mapping.select_account("ninguem") is None


def test_link_e_resolucao_por_subject(mapping):
    mapping.link_account(
        subject="subj-1", access_token="at", refresh_token="rt",
        expires_at=FIXED_NOW + timedelta(hours=1), home_account_id="acc-1",
    )
    session = mapping.get_session("subj-1")
    assert session is not None
    assert session.subject == "subj-1"
    assert session.default_account.account_id == "acc-1"


def test_selecao_entre_contas(mapping):
    mapping.link_account(subject="subj-1", access_token="a1", refresh_token=None,
                         expires_at=None, home_account_id="acc-1", is_default=True)
    mapping.link_account(subject="subj-1", access_token="a2", refresh_token=None,
                         expires_at=None, home_account_id="acc-2", is_default=False)
    assert mapping.select_account("subj-1").account_id == "acc-1"  # default
    assert mapping.select_account("subj-1", "acc-2").account_id == "acc-2"  # específica


def test_mark_expired_forca_reauth(mapping):
    mapping.link_account(subject="subj-1", access_token="a", refresh_token=None,
                         expires_at=None, home_account_id="acc-1")
    mapping.mark_expired("subj-1")
    assert mapping.get_session("subj-1") is None


def test_reuso_de_sessao_existente(mapping):
    sid1 = mapping.ensure_session("subj-1")
    sid2 = mapping.ensure_session("subj-1")
    assert sid1 == sid2  # não cria nova sessão se já existe uma ativa
