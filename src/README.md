# src — código técnico

Área do **código técnico** do servidor MCP (a desenvolver). Mantida separada dos
documentos funcionais, que vivem em [`../docs`](../docs).

## Estrutura prevista

- `prompts/` — assets de runtime carregados pelo servidor (ex.:
  [`assistant-playbook.md`](prompts/assistant-playbook.md), as instruções de orquestração que
  o assistente consulta em execução).
- _(a criar)_ código do servidor MCP, integração Microsoft Graph, fluxo OAuth, camada de
  tokens, ferramentas `*_prepare`/`*_confirm`, etc.

> Pré-requisito antes de codificar: PoC Fase 0 — ver
> [análise funcional v1.1 §2.4 e §10](../docs/analise-funcional-v1.1.md).
