# Análise Funcional — Servidor MCP Office 365
## Versão 1.0 (funcional)

**Cliente:** Mobiweb
**Produto:** Servidor MCP de integração com Microsoft Office 365 (Microsoft Graph API)
**Data:** 2026-06-01
**Estado:** Baseline funcional — superado pela [v1.1](analise-funcional-v1.1.md), que acrescenta a arquitetura de identidade e as decisões de risco. Mantido para histórico.

> **Nota:** esta versão amadureceu o âmbito **funcional** mas tratou a arquitetura de identidade/autenticação de forma superficial. Foi revista após análise crítica independente. Ver [v1.1](analise-funcional-v1.1.md) para a versão a implementar.

---

## 1. Introdução e Objetivo

Servidor MCP (Model Context Protocol) **remoto, self-hosted**, que expõe operações de Office 365 — Email, Calendário, Teams e Ficheiros — como ferramentas utilizáveis a partir do Claude Desktop e Mobile. Permite a utilizadores técnicos e não técnicos executar operações de leitura e escrita sobre as suas caixas O365 por linguagem natural, com **aprovação humana obrigatória antes de qualquer escrita**.

### 1.1 Princípios de design
- **Menor privilégio:** delegated permissions, scopes mínimos por funcionalidade.
- **Consentimento explícito:** todas as ações com efeito externo requerem aprovação humana.
- **Auditabilidade:** todas as operações de escrita são registadas.
- **RGPD/GDPR by design:** minimização de dados e retenção limitada.

## 2. Contexto do projeto

- Equipa pequena: 20-30 utilizadores, técnicos e não técnicos, em Claude Desktop/Mobile.
- Self-hosted numa VPS que já corre outras aplicações.
- Tenant Azure AD/Entra ID já configurado, com acesso admin completo.
- Multi-conta desejável se tecnicamente viável.
- Compliance: dados na UE, aprovação humana antes de escrita.

## 3. Arquitetura (resumo)

| Componente | Decisão |
|---|---|
| Tipo de servidor | MCP remoto, self-hosted em VPS na UE |
| Clientes | Claude Desktop e Mobile |
| Autenticação | OAuth 2.0 (Authorization Code + PKCE) contra Entra ID; refresh tokens geridos pelo servidor |
| Permissões Graph | Delegated (atua em nome do utilizador) |
| API alvo | Microsoft Graph v1.0 |
| Domínio | `mw-mcp-office365.mobiweb.pt` (Cloudflare → VPS) |
| Aprovação | Inline na conversa Claude (human-in-the-loop) |

> Convenção transversal de escrita: **[AC-WRITE]** — *o assistente apresenta um resumo da ação e só executa após confirmação explícita do utilizador*; **[AC-AUDIT]** — *a operação fica registada no log de auditoria*.

## 4. Módulos e User Stories

### Módulo 1 — Email (Outlook)

| ID | User Story | Critérios de aceitação | Notas técnicas |
|----|-----------|------------------------|----------------|
| US-1.1 | **Como** utilizador, **quero** pesquisar emails por remetente, assunto, data ou palavras-chave, **para** encontrar mensagens rapidamente. | Filtros por `from`, `subject`, intervalo de datas, texto livre e pasta. Devolve assunto, remetente, data, snippet, indicação de anexos. Suporta paginação. | `GET /me/messages`, `$search`/`$filter`, `$top`, `$select`. `Mail.Read` |
| US-1.2 | **Como** utilizador, **quero** ler o conteúdo completo de um email. | Devolve corpo + metadados. | `GET /me/messages/{id}`. `Mail.Read` |
| US-1.3 | **Como** utilizador, **quero** enviar um email novo. | to/cc/bcc, assunto, corpo (texto/HTML), anexos. [AC-WRITE], [AC-AUDIT]. | `POST /me/sendMail`. `Mail.Send` |
| US-1.4 | **Como** utilizador, **quero** responder/responder a todos/reencaminhar. | Distingue reply de replyAll; mantém o thread. [AC-WRITE], [AC-AUDIT]. | `POST /me/messages/{id}/reply`, `/replyAll`, `/forward`. `Mail.Send` |
| US-1.5 | **Como** utilizador, **quero** listar e descarregar anexos de um email. | Suporta tipos comuns; respeita limites Graph. | `GET /me/messages/{id}/attachments`. `Mail.Read` |
| US-1.6 | **Como** utilizador, **quero** enviar emails com anexos. | Inline até ~3 MB; *upload session* para maiores. [AC-WRITE], [AC-AUDIT]. | `Mail.Send`, `Mail.ReadWrite` |
| US-1.7 | **Como** utilizador, **quero** arquivar/mover um email. | Lista pastas; resolve nomes→IDs. [AC-WRITE], [AC-AUDIT]. | `POST /me/messages/{id}/move`; `GET /me/mailFolders`. `Mail.ReadWrite` |
| US-1.8 | **Como** utilizador, **quero** apagar um email. | Soft delete por defeito; permanente exige confirmação reforçada. [AC-WRITE], [AC-AUDIT]. | `DELETE /me/messages/{id}`. `Mail.ReadWrite` |

**Scopes:** `Mail.Read`, `Mail.Send`, `Mail.ReadWrite`.

### Módulo 2 — Calendário

| ID | User Story | Critérios de aceitação | Notas técnicas |
|----|-----------|------------------------|----------------|
| US-2.1 | **Como** utilizador, **quero** consultar eventos num intervalo. | Filtro por datas; devolve título, hora, local, organizador, participantes, estado da minha resposta. Respeita fuso. | `GET /me/calendarView`. `Calendars.Read` |
| US-2.2 | **Como** utilizador, **quero** verificar disponibilidade. | Indica conflitos; opcional free/busy de participantes. | `POST /me/calendar/getSchedule`. `Calendars.Read` |
| US-2.3 | **Como** utilizador, **quero** criar uma reunião. | Título, hora, duração, participantes, local/online Teams. [AC-WRITE], [AC-AUDIT]. | `POST /me/events` (`isOnlineMeeting`). `Calendars.ReadWrite` |
| US-2.4 | **Como** utilizador, **quero** editar uma reunião. | Atualiza hora/participantes/assunto; notifica participantes. [AC-WRITE], [AC-AUDIT]. | `PATCH /me/events/{id}`. `Calendars.ReadWrite` |
| US-2.5 | **Como** organizador, **quero** cancelar uma reunião. | Notifica participantes; permite mensagem. [AC-WRITE], [AC-AUDIT]. | `POST /me/events/{id}/cancel`. `Calendars.ReadWrite` |
| US-2.6 | **Como** utilizador, **quero** aceitar/recusar/marcar como tentativo um convite. | Três estados; comentário opcional. [AC-WRITE], [AC-AUDIT]. | `POST /me/events/{id}/accept`\|`/decline`\|`/tentativelyAccept`. `Calendars.ReadWrite` |

**Scopes:** `Calendars.Read`, `Calendars.ReadWrite`.

### Módulo 3 — Teams (apenas chats 1:1 e de grupo)

> **Âmbito v1:** apenas chats 1:1 e de grupo. **Canais de equipas fora de âmbito.**

| ID | User Story | Critérios de aceitação | Notas técnicas |
|----|-----------|------------------------|----------------|
| US-3.1 | **Como** utilizador, **quero** listar os meus chats (1:1 e grupo). | Lista chats a que pertenço. | `GET /me/chats`. `Chat.Read` |
| US-3.2 | **Como** utilizador, **quero** ler as mensagens de um chat. | Remetente, conteúdo, timestamp. | `GET /me/chats/{id}/messages`. `Chat.Read` |
| US-3.3 | **Como** utilizador, **quero** enviar/responder numa mensagem de chat. | Novo chat 1:1, grupo existente, resposta em thread. [AC-WRITE], [AC-AUDIT]. | `POST /chats/{id}/messages`. `Chat.ReadWrite` |

**Scopes:** `Chat.Read`, `Chat.ReadWrite`.

### Módulo 4 — Ficheiros (OneDrive / SharePoint)

| ID | User Story | Critérios de aceitação | Notas técnicas |
|----|-----------|------------------------|----------------|
| US-4.1 | **Como** utilizador, **quero** pesquisar ficheiros por nome ou conteúdo. | Abrange OneDrive e sites SharePoint a que tenho acesso. | `GET /me/drive/root/search`. `Files.Read`/`Sites.Read.All` |
| US-4.2 | **Como** utilizador, **quero** listar o conteúdo de uma pasta. | | `GET /me/drive/items/{id}/children` |
| US-4.3 | **Como** utilizador, **quero** descarregar/ler um ficheiro. | | `GET /me/drive/items/{id}/content` |
| US-4.4 | **Como** utilizador, **quero** carregar um ficheiro. | *Upload session* para ficheiros grandes. [AC-WRITE], [AC-AUDIT]. | `PUT .../content`, `createUploadSession`. `Files.ReadWrite` |
| US-4.5 | **Como** utilizador, **quero** mover/renomear/eliminar um ficheiro. | [AC-WRITE], [AC-AUDIT]. | `PATCH`/`DELETE /me/drive/items/{id}`. `Files.ReadWrite` |

**Scopes:** `Files.ReadWrite`, `Sites.Read.All`.

### Scopes base
`User.Read`, `offline_access`, `openid`, `profile`.

## 5. Aprovação e Auditoria

- **Aprovação inline obrigatória** para todas as operações de escrita/envio.
- **Log de auditoria:** timestamp UTC, utilizador (UPN/ID), ferramenta, tipo de operação, alvo (ID, sem conteúdo integral), resultado, indicação de aprovação. Retenção **12 meses** com purga automática. Metadados apenas, acesso restrito, cifrado em repouso.

## 6. Fora de âmbito (v1)

- Criação de tarefas (Microsoft To Do / Planner) — diferido.
- Mensagens em canais de equipas do Teams — diferido.
- Integrações fora do Office 365.
- Application permissions / acesso a caixas de terceiros sem autorização do próprio.

## 7. Faseamento sugerido

1. **Fase 0 — Piloto técnico:** OAuth remoto + Claude Desktop/Mobile + 1 conta + leitura de email.
2. **Fase 1 — Email** (incl. anexos) + auditoria + aprovação.
3. **Fase 2 — Calendário.**
4. **Fase 3 — Teams** (chats).
5. **Fase 4 — Ficheiros + multi-conta.**
6. **Futuro — Tarefas, canais Teams.**

---

*Documento substituído pela [Análise Funcional v1.1](analise-funcional-v1.1.md).*
