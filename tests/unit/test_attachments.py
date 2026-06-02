"""Unit — extração de texto de anexos (graph/attachments.py).

Cobre PDF (com fixture real), texto simples, tipo não suportado e base64 inválido. Garante
que o servidor devolve texto legível em vez de base64 cru (regressão do problema reportado
com a fatura em PDF).
"""

from __future__ import annotations

import base64
from pathlib import Path

from mcp_o365.graph.attachments import extract_attachment_text

_FIXTURE = Path(__file__).resolve().parents[1] / "fixtures" / "sample_invoice.pdf"


def _b64(data: bytes) -> str:
    return base64.b64encode(data).decode()


def test_extrai_texto_de_pdf_real():
    pdf = _FIXTURE.read_bytes()
    res = extract_attachment_text(
        name="fatura.pdf", content_type="application/pdf", content_bytes_b64=_b64(pdf)
    )
    assert res["extractable"] is True
    assert "1234.56 EUR" in res["text"]
    assert res["pages"] == 1
    assert res["truncated"] is False


def test_extrai_texto_simples_por_content_type():
    res = extract_attachment_text(
        name="nota.txt", content_type="text/plain",
        content_bytes_b64=_b64("Olá mundo €".encode()),
    )
    assert res["extractable"] is True
    assert res["text"] == "Olá mundo €"


def test_deteta_pdf_pelo_nome_mesmo_sem_content_type():
    pdf = _FIXTURE.read_bytes()
    res = extract_attachment_text(
        name="documento.PDF", content_type=None, content_bytes_b64=_b64(pdf)
    )
    assert res["extractable"] is True
    assert "1234.56" in res["text"]


def test_tipo_nao_suportado_nao_extrai_e_sugere_include_bytes():
    res = extract_attachment_text(
        name="imagem.png", content_type="image/png", content_bytes_b64=_b64(b"\x89PNG..."),
    )
    assert res["extractable"] is False
    assert "include_bytes" in res["reason"]


def test_pdf_invalido_nao_rebenta():
    res = extract_attachment_text(
        name="corrompido.pdf", content_type="application/pdf",
        content_bytes_b64=_b64(b"isto nao e um pdf"),
    )
    assert res["extractable"] is False
    assert res.get("reason")


def test_base64_invalido_e_vazio():
    assert extract_attachment_text(
        name="x.txt", content_type="text/plain", content_bytes_b64=None
    )["extractable"] is False
    assert extract_attachment_text(
        name="x.txt", content_type="text/plain", content_bytes_b64="!!!nao-base64!!!"
    )["extractable"] is False
