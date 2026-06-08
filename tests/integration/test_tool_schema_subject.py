"""Regressão — o assunto de emails/eventos é exposto às tools MCP como `subject`.

Bug reportado: "o servidor está a ignorar o campo subject no prepare". A causa não era uma
limitação do MCP, mas um desalinhamento de nomes: as tools de escrita expunham o assunto como
`subject_line`. Como o principal autenticado (o `subject` interno do Plano A) vem do token e
NUNCA é um parâmetro de tool, não havia colisão na fronteira — pelo que o parâmetro passou a
chamar-se `subject`. Um cliente que envia `subject` (o nome natural, igual ao do Graph) deixa
de ver o campo descartado silenciosamente.

Estes testes fixam o contrato do schema das tools para evitar a regressão.
"""

from __future__ import annotations

import asyncio

import pytest

from mcp_o365.app import build_components

# Tools de escrita que aceitam um assunto e os parâmetros esperados no schema MCP.
_TOOLS_WITH_SUBJECT = ("email_send_prepare", "calendar_create_prepare", "calendar_update_prepare")


@pytest.fixture
def tools_by_name(config):
    comp = build_components(config)
    # build_server regista as tools; list_tools devolve mcp.types.Tool com inputSchema (JSON).
    tools = asyncio.run(comp["server"].list_tools())
    return {t.name: t for t in tools}


@pytest.mark.parametrize("tool_name", _TOOLS_WITH_SUBJECT)
def test_subject_is_exposed_and_subject_line_is_not(tools_by_name, tool_name):
    tool = tools_by_name[tool_name]
    props = tool.inputSchema.get("properties", {})
    assert "subject" in props, f"{tool_name} deve expor o assunto como 'subject'"
    assert "subject_line" not in props, (
        f"{tool_name} não deve expor 'subject_line' (nome natural/Graph é 'subject')"
    )
