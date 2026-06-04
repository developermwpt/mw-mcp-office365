"""Extração de texto de anexos — para o modelo ler o conteúdo sem manipular base64.

Devolver o anexo como base64 cru obriga o cliente (Claude) a descodificar e a tentar
extrair texto à mão, o que é frágil (sobretudo em PDFs). Aqui, no servidor, extraímos o
texto dos tipos comuns (PDF, texto) e devolvemos texto legível; os bytes só seguem quando
explicitamente pedidos (`include_bytes`).

O conteúdo de anexos é, tal como os corpos de email, **NÃO-confiável** (prompt injection):
o texto extraído é devolvido marcado como tal e nunca deve ser tratado como instruções.
"""

from __future__ import annotations

import base64
import io
import logging
import xml.etree.ElementTree as ET
import zipfile

logger = logging.getLogger("mcp_o365.graph.attachments")

# Namespaces OOXML (Word/PowerPoint).
_W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

# Limite de caracteres devolvidos ao modelo (evita despejar documentos enormes no contexto).
_MAX_TEXT_CHARS = 50_000

_TEXT_CONTENT_TYPES = frozenset(
    {"application/json", "application/xml", "text/xml", "application/csv"}
)
_TEXT_SUFFIXES = (".txt", ".csv", ".md", ".json", ".xml", ".log")


def _is_pdf(name: str | None, content_type: str | None) -> bool:
    return (content_type or "").lower().startswith("application/pdf") or (
        (name or "").lower().endswith(".pdf")
    )


def _is_docx(name: str | None, content_type: str | None) -> bool:
    ct = (content_type or "").lower()
    return (
        "wordprocessingml.document" in ct
        or (name or "").lower().endswith(".docx")
    )


def _is_pptx(name: str | None, content_type: str | None) -> bool:
    ct = (content_type or "").lower()
    return (
        "presentationml.presentation" in ct
        or (name or "").lower().endswith(".pptx")
    )


def _is_legacy_office(name: str | None) -> bool:
    return (name or "").lower().endswith((".doc", ".ppt", ".xls"))


def _is_text(name: str | None, content_type: str | None) -> bool:
    ct = (content_type or "").lower().split(";")[0].strip()
    return (
        ct.startswith("text/")
        or ct in _TEXT_CONTENT_TYPES
        or (name or "").lower().endswith(_TEXT_SUFFIXES)
    )


def _truncate(text: str) -> dict:
    truncated = len(text) > _MAX_TEXT_CHARS
    return {
        "extractable": True,
        "text": text[:_MAX_TEXT_CHARS],
        "truncated": truncated,
    }


def _extract_pdf(raw: bytes) -> dict:
    try:
        from pypdf import PdfReader
    except Exception:  # noqa: BLE001 — dependência opcional ausente
        return {"extractable": False, "reason": "extração de PDF indisponível no servidor."}
    try:
        reader = PdfReader(io.BytesIO(raw))
        pages = len(reader.pages)
        text = "\n".join((p.extract_text() or "") for p in reader.pages).strip()
    except Exception:  # noqa: BLE001 — PDF malformado/cifrado não deve rebentar a leitura
        logger.warning(
            "falha a extrair texto do PDF",
            extra={"fields": {"event": "attachment_extract_error", "kind": "pdf"}},
        )
        return {"extractable": False, "reason": "não foi possível ler o PDF."}
    out = _truncate(text)
    out["pages"] = pages
    if not text:
        # PDF digitalizado (imagem) — sem camada de texto. OCR está fora de âmbito.
        out["extractable"] = False
        out["reason"] = "PDF sem texto extraível (provavelmente digitalizado/imagem)."
    return out


def _extract_docx(raw: bytes) -> dict:
    """Extrai texto de um .docx (OOXML = ZIP com XML). Só biblioteca-padrão."""
    try:
        with zipfile.ZipFile(io.BytesIO(raw)) as zf:
            xml = zf.read("word/document.xml")
    except (zipfile.BadZipFile, KeyError, OSError):
        logger.warning(
            "falha a abrir o .docx",
            extra={"fields": {"event": "attachment_extract_error", "kind": "docx"}},
        )
        return {"extractable": False, "reason": "não foi possível ler o .docx."}
    try:
        root = ET.fromstring(xml)
    except ET.ParseError:
        return {"extractable": False, "reason": "não foi possível ler o .docx (XML)."}

    # Cada parágrafo (w:p) → junta o texto dos seus runs (w:t); parágrafos por linha.
    paragraphs = []
    for para in root.iter(f"{_W_NS}p"):
        runs = [node.text or "" for node in para.iter(f"{_W_NS}t")]
        paragraphs.append("".join(runs))
    text = "\n".join(paragraphs).strip()
    if not text:
        return {"extractable": False, "reason": ".docx sem texto extraível."}
    return _truncate(text)


def _extract_pptx(raw: bytes) -> dict:
    """Extrai texto de um .pptx (OOXML), slide a slide. Só biblioteca-padrão."""
    try:
        with zipfile.ZipFile(io.BytesIO(raw)) as zf:
            slides = sorted(
                n for n in zf.namelist()
                if n.startswith("ppt/slides/slide") and n.endswith(".xml")
            )
            chunks = []
            for slide_name in slides:
                try:
                    root = ET.fromstring(zf.read(slide_name))
                except ET.ParseError:
                    continue
                texts = [node.text or "" for node in root.iter(f"{_A_NS}t")]
                slide_text = "\n".join(t for t in texts if t.strip())
                if slide_text:
                    chunks.append(slide_text)
    except (zipfile.BadZipFile, OSError):
        logger.warning(
            "falha a abrir o .pptx",
            extra={"fields": {"event": "attachment_extract_error", "kind": "pptx"}},
        )
        return {"extractable": False, "reason": "não foi possível ler o .pptx."}
    text = "\n\n".join(chunks).strip()
    if not text:
        return {"extractable": False, "reason": ".pptx sem texto extraível."}
    return _truncate(text)


def extract_attachment_text(
    *, name: str | None, content_type: str | None, content_bytes_b64: str | None
) -> dict:
    """Extrai texto de um anexo, se o tipo for suportado.

    Devolve `{"extractable": True, "text": str, "truncated": bool[, "pages": int]}` ou
    `{"extractable": False, "reason": str}`. Nunca levanta — o caminho de leitura de email
    não pode falhar por causa de um anexo problemático.
    """
    if not content_bytes_b64:
        return {"extractable": False, "reason": "sem conteúdo (contentBytes vazio)."}
    try:
        raw = base64.b64decode(content_bytes_b64)
    except Exception:  # noqa: BLE001
        return {"extractable": False, "reason": "contentBytes inválido (base64)."}

    if _is_pdf(name, content_type):
        return _extract_pdf(raw)
    if _is_docx(name, content_type):
        return _extract_docx(raw)
    if _is_pptx(name, content_type):
        return _extract_pptx(raw)
    if _is_text(name, content_type):
        return _truncate(raw.decode("utf-8", errors="replace"))
    if _is_legacy_office(name):
        return {
            "extractable": False,
            "reason": (
                "formato Office legado (.doc/.ppt/.xls) não suportado para extração; "
                "peça o ficheiro em formato moderno (.docx/.pptx/.xlsx)."
            ),
        }
    return {
        "extractable": False,
        "reason": (
            f"tipo '{content_type or 'desconhecido'}' não suportado para extração de "
            "texto; use include_bytes=True para obter os bytes em base64."
        ),
    }
