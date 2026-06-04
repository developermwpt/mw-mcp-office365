"""Unit — extração de texto de anexos (graph/attachments.py).

Cobre PDF (com fixture real), texto simples, tipo não suportado e base64 inválido. Garante
que o servidor devolve texto legível em vez de base64 cru (regressão do problema reportado
com a fatura em PDF).
"""

from __future__ import annotations

import base64
import io
import zipfile
from pathlib import Path

from mcp_o365.graph.attachments import extract_attachment_text

_FIXTURE = Path(__file__).resolve().parents[1] / "fixtures" / "sample_invoice.pdf"

_DOCX_CT = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
)
_PPTX_CT = (
    "application/vnd.openxmlformats-officedocument.presentationml.presentation"
)


def _b64(data: bytes) -> str:
    return base64.b64encode(data).decode()


def _make_docx(paragraphs: list[str]) -> bytes:
    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    body = "".join(
        f"<w:p><w:r><w:t>{p}</w:t></w:r></w:p>" for p in paragraphs
    )
    xml = (
        f'<?xml version="1.0"?><w:document xmlns:w="{w}"><w:body>{body}</w:body>'
        "</w:document>"
    )
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("word/document.xml", xml)
    return buf.getvalue()


def _make_pptx(slides: list[list[str]]) -> bytes:
    a = "http://schemas.openxmlformats.org/drawingml/2006/main"
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        for i, texts in enumerate(slides, start=1):
            runs = "".join(f"<a:t>{t}</a:t>" for t in texts)
            xml = f'<?xml version="1.0"?><sld xmlns:a="{a}">{runs}</sld>'
            zf.writestr(f"ppt/slides/slide{i}.xml", xml)
    return buf.getvalue()


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


def test_extrai_texto_de_docx():
    raw = _make_docx(["Proposta comercial Mobiweb", "Valor total: 4500 EUR"])
    res = extract_attachment_text(
        name="proposta.docx", content_type=_DOCX_CT, content_bytes_b64=_b64(raw)
    )
    assert res["extractable"] is True
    assert "Proposta comercial Mobiweb" in res["text"]
    assert "4500 EUR" in res["text"]
    # parágrafos separados por nova linha
    assert "\n" in res["text"]


def test_deteta_docx_pelo_nome_sem_content_type():
    raw = _make_docx(["Conteúdo do documento"])
    res = extract_attachment_text(
        name="DOCUMENTO.DOCX", content_type=None, content_bytes_b64=_b64(raw)
    )
    assert res["extractable"] is True
    assert "Conteúdo do documento" in res["text"]


def test_docx_corrompido_nao_rebenta():
    res = extract_attachment_text(
        name="mau.docx", content_type=_DOCX_CT, content_bytes_b64=_b64(b"isto nao e zip")
    )
    assert res["extractable"] is False
    assert res.get("reason")


def test_extrai_texto_de_pptx():
    raw = _make_pptx([["Título da apresentação"], ["Slide 2", "ponto chave"]])
    res = extract_attachment_text(
        name="deck.pptx", content_type=_PPTX_CT, content_bytes_b64=_b64(raw)
    )
    assert res["extractable"] is True
    assert "Título da apresentação" in res["text"]
    assert "ponto chave" in res["text"]


def test_formato_office_legado_da_mensagem_clara():
    res = extract_attachment_text(
        name="antigo.doc", content_type="application/msword",
        content_bytes_b64=_b64(b"\xd0\xcf\x11\xe0"),
    )
    assert res["extractable"] is False
    assert ".docx" in res["reason"]


def test_base64_invalido_e_vazio():
    assert extract_attachment_text(
        name="x.txt", content_type="text/plain", content_bytes_b64=None
    )["extractable"] is False
    assert extract_attachment_text(
        name="x.txt", content_type="text/plain", content_bytes_b64="!!!nao-base64!!!"
    )["extractable"] is False
