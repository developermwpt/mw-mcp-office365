"""T3 — Cifra de tokens.

Interface `Cipher` (abstrata) + implementação local AES-256-GCM. Toda a cifra de tokens
passa por aqui; quando se migrar para um KMS/Key Vault (v1.1 §4), basta uma nova
implementação de `Cipher` sem tocar no resto do código.
"""

from __future__ import annotations

import os
from abc import ABC, abstractmethod

from cryptography.hazmat.primitives.ciphers.aead import AESGCM

_NONCE_LEN = 12  # tamanho recomendado para AES-GCM


class Cipher(ABC):
    """Contrato de cifra simétrica autenticada (AEAD)."""

    @abstractmethod
    def encrypt(self, plaintext: bytes) -> bytes:
        ...

    @abstractmethod
    def decrypt(self, ciphertext: bytes) -> bytes:
        ...

    def encrypt_str(self, value: str) -> bytes:
        return self.encrypt(value.encode("utf-8"))

    def decrypt_str(self, blob: bytes) -> str:
        return self.decrypt(blob).decode("utf-8")


class LocalAesGcmCipher(Cipher):
    """AES-256-GCM com chave local. Nonce aleatório de 12 bytes prefixado ao ciphertext.

    Formato em repouso: ``nonce(12) || ciphertext+tag``.
    """

    def __init__(self, key: bytes) -> None:
        if len(key) != 32:
            raise ValueError("A chave AES-256 tem de ter 32 bytes.")
        self._aes = AESGCM(key)

    def encrypt(self, plaintext: bytes) -> bytes:
        nonce = os.urandom(_NONCE_LEN)
        ct = self._aes.encrypt(nonce, plaintext, None)
        return nonce + ct

    def decrypt(self, ciphertext: bytes) -> bytes:
        if len(ciphertext) < _NONCE_LEN + 1:
            raise ValueError("Ciphertext demasiado curto.")
        nonce, ct = ciphertext[:_NONCE_LEN], ciphertext[_NONCE_LEN:]
        return self._aes.decrypt(nonce, ct, None)  # InvalidTag se a chave/dados não baterem
