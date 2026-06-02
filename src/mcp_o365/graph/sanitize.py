"""Sanitização de HTML de emails — mitigação de prompt injection (v1.1 §4).

O corpo de um email é conteúdo NÃO-confiável: pode conter instruções escondidas dirigidas
ao assistente (prompt injection), scripts, ou conteúdo invisível. Antes de qualquer corpo
HTML ser devolvido ao modelo, passa por aqui.

Usa apenas `html.parser.HTMLParser` da biblioteca padrão (sem bleach/lxml). Remove:
- `<script>` e `<style>` (incluindo o seu conteúdo);
- comentários HTML;
- atributos de event handler (`on*`) e URIs `javascript:`;
- conteúdo óbvio invisível (`display:none` / `visibility:hidden`).

Ambas as funções são robustas a HTML malformado — o `HTMLParser` da stdlib é tolerante e
nunca levanta em entrada normal; ainda assim apanhamos qualquer exceção e devolvemos o que
foi acumulado, para nunca rebentar o caminho de leitura de email.
"""

from __future__ import annotations

import re
from html import escape
from html.parser import HTMLParser

# Tags cujo conteúdo é descartado por completo (não só as tags).
_DROP_CONTENT_TAGS = frozenset({"script", "style"})
# Tags vazias (sem fecho) comuns em HTML.
_VOID_TAGS = frozenset(
    {"br", "img", "hr", "input", "meta", "link", "area", "base", "col", "embed",
     "source", "track", "wbr"}
)
# Atributos cujo valor pode esconder/injetar — sempre removidos.
_BLOCKED_ATTRS = frozenset({"style", "srcset"})

_HIDDEN_RE = re.compile(
    r"display\s*:\s*none|visibility\s*:\s*hidden", re.IGNORECASE
)
_JS_URI_RE = re.compile(r"^\s*javascript:", re.IGNORECASE)
_URI_ATTRS = frozenset({"href", "src", "action", "formaction", "background"})


class _SanitizingParser(HTMLParser):
    """Reconstrói HTML limpo, descartando scripts/estilos e atributos perigosos."""

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self._out: list[str] = []
        # Profundidade dentro de blocos cujo conteúdo deve ser totalmente ignorado.
        self._drop_depth = 0

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag in _DROP_CONTENT_TAGS:
            self._drop_depth += 1
            return
        if self._drop_depth:
            return
        if self._is_hidden(attrs):
            # Conteúdo invisível: trata como bloco a descartar até ao fecho.
            self._drop_depth += 1
            return
        safe = self._clean_attrs(attrs)
        if tag in _VOID_TAGS:
            self._out.append(f"<{tag}{safe}>")
        else:
            self._out.append(f"<{tag}{safe}>")

    def handle_startendtag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag in _DROP_CONTENT_TAGS or self._drop_depth:
            return
        if self._is_hidden(attrs):
            return
        safe = self._clean_attrs(attrs)
        self._out.append(f"<{tag}{safe}/>")

    def handle_endtag(self, tag: str) -> None:
        if tag in _DROP_CONTENT_TAGS:
            if self._drop_depth:
                self._drop_depth -= 1
            return
        if self._drop_depth:
            self._drop_depth -= 1
            return
        if tag not in _VOID_TAGS:
            self._out.append(f"</{tag}>")

    def handle_data(self, data: str) -> None:
        if self._drop_depth:
            return
        self._out.append(escape(data, quote=False))

    def handle_comment(self, data: str) -> None:
        # Comentários são descartados (podem esconder instruções).
        return

    @staticmethod
    def _is_hidden(attrs: list[tuple[str, str | None]]) -> bool:
        for name, value in attrs:
            if name.lower() == "style" and value and _HIDDEN_RE.search(value):
                return True
            if name.lower() == "hidden":
                return True
        return False

    @staticmethod
    def _clean_attrs(attrs: list[tuple[str, str | None]]) -> str:
        parts: list[str] = []
        for name, value in attrs:
            low = name.lower()
            if low.startswith("on"):  # event handlers (onclick, onerror, ...)
                continue
            if low in _BLOCKED_ATTRS:
                continue
            if value is None:
                parts.append(f" {low}")
                continue
            if low in _URI_ATTRS and _JS_URI_RE.match(value):
                continue  # remove javascript: URIs
            parts.append(f' {low}="{escape(value, quote=True)}"')
        return "".join(parts)

    def result(self) -> str:
        return "".join(self._out)


class _TextParser(HTMLParser):
    """Extrai apenas o texto visível, descartando scripts/estilos e tags."""

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self._chunks: list[str] = []
        self._drop_depth = 0

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag in _DROP_CONTENT_TAGS:
            self._drop_depth += 1

    def handle_endtag(self, tag: str) -> None:
        if tag in _DROP_CONTENT_TAGS and self._drop_depth:
            self._drop_depth -= 1

    def handle_data(self, data: str) -> None:
        if not self._drop_depth and data.strip():
            self._chunks.append(data)

    def handle_comment(self, data: str) -> None:
        return

    def result(self) -> str:
        return " ".join(" ".join(self._chunks).split())


def sanitize_html(html: str) -> str:
    """Devolve HTML limpo: sem scripts/estilos, event handlers, `javascript:` ou ocultos."""
    if not html:
        return ""
    parser = _SanitizingParser()
    try:
        parser.feed(html)
        parser.close()
    except Exception:  # noqa: BLE001 — nunca rebentar o caminho de leitura
        pass
    return parser.result()


def html_to_text(html: str) -> str:
    """Versão texto simples (sem tags), útil para resumos. Robusta a HTML malformado."""
    if not html:
        return ""
    parser = _TextParser()
    try:
        parser.feed(html)
        parser.close()
    except Exception:  # noqa: BLE001 — nunca rebentar o caminho de leitura
        pass
    return parser.result()
