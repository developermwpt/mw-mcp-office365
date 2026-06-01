"""Unit — cifra (T3): round-trip, unicidade de nonce, falha com chave errada."""

from __future__ import annotations

import pytest
from cryptography.exceptions import InvalidTag

from mcp_o365.storage.crypto import LocalAesGcmCipher


def test_round_trip():
    c = LocalAesGcmCipher(b"\x01" * 32)
    assert c.decrypt(c.encrypt(b"segredo")) == b"segredo"
    assert c.decrypt_str(c.encrypt_str("olá mundo")) == "olá mundo"


def test_nonce_unico_ciphertext_difere():
    c = LocalAesGcmCipher(b"\x02" * 32)
    a = c.encrypt(b"mesmo texto")
    b = c.encrypt(b"mesmo texto")
    assert a != b  # nonce aleatório por operação
    assert c.decrypt(a) == c.decrypt(b) == b"mesmo texto"


def test_chave_errada_falha():
    c1 = LocalAesGcmCipher(b"\x03" * 32)
    c2 = LocalAesGcmCipher(b"\x04" * 32)
    blob = c1.encrypt(b"x")
    with pytest.raises(InvalidTag):
        c2.decrypt(blob)


def test_chave_tamanho_invalido():
    with pytest.raises(ValueError):
        LocalAesGcmCipher(b"curta")
