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

logger = logging.getLogger("mcp_o365.graph.attachments")

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
    if _is_text(name, content_type):
        return _truncate(raw.decode("utf-8", errors="replace"))
    return {
        "extractable": False,
        "reason": (
            f"tipo '{content_type or 'desconhecido'}' não suportado para extração de "
            "texto; use include_bytes=True para obter os bytes em base64."
        ),
    }
