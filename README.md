# mw-mcp-office365

Servidor MCP (Model Context Protocol) remoto para integração com o Microsoft Office 365
(Email, Calendário, Teams e Ficheiros) via Microsoft Graph, para uso no Claude.

## Documentação

- [`docs/analise-funcional-v1.1.md`](docs/analise-funcional-v1.1.md) — **Análise funcional a
  implementar.** User stories + arquitetura de identidade (duplo OAuth), aprovação server-side
  two-phase, segurança/prompt injection, RGPD e decisões de risco.
- [`docs/analise-funcional-v1.0.md`](docs/analise-funcional-v1.0.md) — Baseline funcional
  (user stories + critérios de aceitação), mantido para histórico.
- [`docs/assistant-playbook.md`](docs/assistant-playbook.md) — Playbook de instruções que o
  assistente consulta em runtime: princípios de operação, guia por ferramenta, orquestração
  de pedidos complexos e interligados, receitas e tabela de referência rápida.

## Estado

Fase de análise/arquitetura. Pré-requisito antes de codificar: PoC Fase 0 (validar connector
remoto no Claude Team Desktop+Mobile e o refresh sob a política de Conditional Access do tenant).
