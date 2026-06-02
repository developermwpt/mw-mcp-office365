# mw-mcp-office365

Servidor MCP (Model Context Protocol) remoto para integração com o Microsoft Office 365
(Email, Calendário, Teams e Ficheiros) via Microsoft Graph, para uso no Claude.

## Estrutura do repositório

```
docs/   → documentos funcionais (análise, user stories) — negócio
src/    → código técnico do servidor MCP (a desenvolver)
          └── prompts/ → assets de runtime carregados pelo servidor
```

## Documentos funcionais — [`docs/`](docs)

- [`docs/analise-funcional-v1.1.md`](docs/analise-funcional-v1.1.md) — **Análise funcional a
  implementar.** User stories + arquitetura de identidade (duplo OAuth), aprovação server-side
  two-phase, segurança/prompt injection, RGPD e decisões de risco.
- [`docs/analise-funcional-v1.0.md`](docs/analise-funcional-v1.0.md) — Baseline funcional
  (user stories + critérios de aceitação), mantido para histórico.

## Código técnico — [`src/`](src)

- [`src/prompts/assistant-playbook.md`](src/prompts/assistant-playbook.md) — Playbook de
  instruções que o servidor carrega em runtime: princípios de operação, guia por ferramenta,
  orquestração de pedidos complexos e interligados, receitas e tabela de referência rápida.

## Estado

**PoC Fase 0 implementada** (T0–T12) — esqueleto do servidor MCP em Python (`mcp` 1.27.2):
dual-plane OAuth (Plano A via SDK: RFC 9728/8414/7591; Plano B via `msal`: authcode+PKCE +
refresh), mapeamento de identidade multi-conta, token store SQLite cifrado, tool read-only
`whoami` e health checks. Ver [`src/`](src) e o
[plano de implementação](docs/poc-fase-0/plano-implementacao.md).

**Fase 1 — Módulo Email implementada** (US-1.1 a US-1.8): pesquisar, ler, listar/descarregar
anexos, enviar, responder/responder-a-todos/reencaminhar, mover e eliminar (soft + permanente
reforçada). Todas as escritas seguem **aprovação em duas fases** (prepare/confirm com token de
uso único, TTL e idempotência), com **auditoria** estruturada só de metadados (`event=audit`,
sem PII) e **sanitização** do corpo HTML dos emails recebidos (mitigação de prompt injection).
Ver o [estado das user stories](docs/fase-1/estado-user-stories.md) e o
[runbook de validação](docs/fase-1/runbook-validacao-email.md).

**113 testes** no total (unit + integração, Graph/Entra mockados) verdes e `ruff` limpo.

**Validada no tenant/VPS reais (2026-06-02):** o gate **G3** (refresh do token Graph **sob a
Conditional Access** em dispositivo gerido) **passou** — o refresh silencioso do servidor
mantém acesso, sem necessidade de exceção de CA. O connector liga no Claude Team Desktop+Mobile
e as US de Email US-1.1–1.5, 1.7 e 1.8 foram exercidas no real com sucesso (ver
[estado das user stories](docs/fase-1/estado-user-stories.md)). Falta apenas exercer no real o
**envio de um anexo >3 MB** (US-1.6; coberto por testes automáticos). Decisões jurídicas
(DPA/DPIA) continuam em aberto para a operação com dados reais a longo prazo.
