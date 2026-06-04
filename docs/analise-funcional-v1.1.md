# Análise Funcional — Servidor MCP Office 365
## Versão 1.1 (funcional + arquitetura de identidade e risco)

**Cliente:** Mobiweb
**Produto:** Servidor MCP de integração com Microsoft Office 365 (Microsoft Graph API)
**Data:** 2026-06-01
**Estado:** Versão a implementar. Substitui a [v1.0](analise-funcional-v1.0.md), que tratava a autenticação de forma superficial.

> **Como esta versão nasceu:** a v1.0 foi sujeita a uma análise crítica independente que identificou que a camada de **arquitetura de identidade de um MCP remoto multi-utilizador** — o ponto mais difícil e arriscado do projeto — estava por resolver, além de riscos de segurança de primeira ordem (prompt injection, idempotência, soberania de dados). Esta v1.1 incorpora essas correções e as decisões de risco tomadas pelo cliente.

---

## 1. Objetivo e princípios

Servidor MCP remoto self-hosted que expõe operações de Office 365 (Email, Calendário, Teams chats, Ficheiros) ao Claude Desktop/Mobile, para ~20-30 utilizadores O365, com leitura e escrita via Microsoft Graph **delegated permissions** e **aprovação humana imposta server-side** antes de qualquer escrita.

Princípios: menor privilégio · consentimento explícito imposto pelo servidor (não apenas narrativo) · auditabilidade · RGPD by design · resistência a prompt injection.

## 2. Arquitetura de Identidade (secção central)

> Esta é a parte que decide a viabilidade do projeto. Coexistem **dois planos OAuth distintos** que têm de ser desenhados explicitamente.

### 2.1 Os dois planos OAuth

- **Plano A — Cliente Claude ↔ Servidor MCP.** O servidor MCP atua como **OAuth Resource Server / Authorization Server** para o cliente Claude. O Claude obtém um access token *para o MCP*. Implica expor:
  - **Protected Resource Metadata (RFC 9728)**
  - **Authorization Server Metadata (RFC 8414)**
  - **Dynamic Client Registration (RFC 7591)** — esperado pelos clientes Claude remotos.
- **Plano B — Servidor MCP ↔ Entra ID.** O servidor é um **client confidencial** que detém os tokens Graph **delegados** de cada utilizador (Authorization Code + PKCE).

### 2.2 Mapeamento de identidade (o elo crítico)

Cada request que chega ao servidor traz o token do **Plano A**. O servidor tem de o resolver, server-side, para a sessão de tokens Graph do **Plano B** do utilizador correto:

```
token MCP (Plano A) → subject/sessão → registo de sessão → refresh/access token Graph cifrado (Plano B)
```

Regras de desenho:
- Isolamento estrito por utilizador na camada de dados (sem acesso cruzado).
- O modelo nasce **multi-conta**: uma sessão MCP → N sessões Graph (uma por conta O365 ligada). Adicionar multi-conta depois implicaria reescrita.
- **Re-autenticação graciosa:** ao detetar `invalid_grant` (refresh token expirado/revogado), marcar a sessão como expirada e pedir novo login ao utilizador dentro do Claude. Nunca falhar silenciosamente uma escrita. "Login uma vez" = "login até o refresh token expirar ou ser revogado".

### 2.3 Decisões de autenticação confirmadas

| Tema | Decisão |
|---|---|
| Fluxo | OAuth 2.0 Authorization Code + PKCE contra Entra ID |
| Permissões | Delegated (limite efetivo = ACLs reais do utilizador) |
| Consentimento | **Admin consent global** (tenant-wide) — UX "login uma vez", sem ecrã por utilizador |
| Plano Claude | **Team** — suporte a connector remoto OAuth a validar em PoC (esp. Mobile) |
| Domínio/TLS | `mw-mcp-office365.mobiweb.pt`, Cloudflare → VPS, Origin Certificate (Full strict). Redirect URI OAuth no domínio final. |

### 2.4 ⚠️ Bloqueador a resolver — Conditional Access

O tenant **exige dispositivo gerido/compliant** (confirmado pelo cliente). Implicação:
- **Login interativo inicial:** ocorre no browser do dispositivo do utilizador (gerido) → **passa**.
- **Refresh silencioso a partir do servidor:** o servidor MCP não é dispositivo gerido → o Entra pode **rejeitar** a renovação silenciosa, degradando ("login uma vez" passa a reautenticação periódica) ou **bloqueando** (com token protection/binding).

**Caminhos (decisão + admin Entra):** (1) exceção de Conditional Access com escopo mínimo — *named location* (IP da VPS) ou a identidade do servidor [mais limpo]; (2) aceitar reautenticação interativa periódica como constraint de design; (3) registar o servidor como dispositivo compliant [frágil].

**Pré-requisito antes de codificar:** PoC Fase 0 que valide (a) o connector remoto no Claude Team Desktop **e** Mobile, e (b) o comportamento real do refresh sob esta política de CA.

## 3. Modelo de aprovação — server-side em duas fases

> A aprovação inline "narrativa" (texto na conversa) **não é um controlo** — um prompt confiante ou um prompt injection contorna-a. O ponto de imposição está **no servidor**.

Toda a operação de escrita divide-se em:
1. **`*_prepare`** — valida, monta o payload, devolve um **resumo + token de confirmação** de uso único com TTL. **Não executa nada no Graph.**
2. **`*_confirm`** — só executa se receber um token de confirmação válido, não expirado e não usado.

O token de confirmação serve também de **idempotency key** (neutraliza retries do transporte e re-chamadas do LLM). Operações de leitura/pesquisa não exigem aprovação.

## 4. Segurança

- **Prompt injection (maior risco do projeto):** conteúdo lido (emails, mensagens Teams, ficheiros) é **não-confiável**. O assistente nunca executa instruções vindas de dentro do conteúdo lido; só age por intenção direta do utilizador. Sanitização de HTML na leitura (remover scripts, conteúdo invisível). Escritas só acionáveis por intenção do utilizador, com resumo real no `confirm`. *Mitiga, não elimina — risco residual aceite pelo cliente.*
- **Idempotência:** idempotency keys por operação (geradas no `prepare`); respeitar `Retry-After`; retry só em operações guardadas por chave.
- **Armazenamento de tokens:** **on-VPS** (decisão de custo do cliente). Risco aceite: comprometer a VPS = expor 20-30 mailboxes = possível notificação à CNPD. Mitigações: utilizador de SO dedicado/contentor isolado, permissões restritas, chave fora da raiz web, camada de cifra abstraída para migrar a Key Vault depois.
- **Rate limits Graph:** por mailbox **e** por app/tenant; backoff exponencial + monitorização da taxa de 429.

## 5. Scopes Microsoft Graph (delegated)

| Módulo | Scopes | Nota |
|---|---|---|
| Email | `Mail.Read`, `Mail.Send`, `Mail.ReadWrite` | Inclui anexos (upload sessions) |
| Calendário | `Calendars.Read`, `Calendars.ReadWrite` | |
| Teams (chats) | `Chat.Read`, `Chat.ReadWrite` | Sem scopes de canal |
| Ficheiros | `Files.ReadWrite`, `Sites.Read.All` | `Files.ReadWrite` (próprio drive); `Sites.Read.All` *delegado* limitado pelas ACLs do utilizador. **Não** `Sites.Selected` (é app-only, incompatível com fluxo delegado). |
| Base | `User.Read`, `offline_access`, `openid`, `profile` | |

## 6. Módulos e User Stories

Idênticos à [v1.0 §4](analise-funcional-v1.0.md), com a alteração transversal: cada **[AC-WRITE]** passa a significar **par `prepare`/`confirm` server-side** (não aprovação narrativa). Resumo:

- **Email:** pesquisar, ler, enviar, responder/responder-a-todos/reencaminhar, anexos (receber/enviar), arquivar/mover, apagar. *Pesquisa por período com paginação consciente:* período ≤ 24h devolve **todos** os emails automaticamente; período > 24h com mais resultados do que cabe numa página **pergunta** ao utilizador (todos vs primeiros N) antes de paginar — ver [estado-user-stories §US-1.1](fase-1/estado-user-stories.md). *Extração de anexos server-side:* o conteúdo de anexos é **extraído para texto no próprio servidor** (PDF, Word `.docx`, PowerPoint `.pptx` e ficheiros de texto) e devolvido em `extracted_text` pronto a ler — o modelo **não** recebe nem descodifica base64 (mais rápido e fiável). Os bytes em base64 só seguem a pedido explícito (`include_bytes`), como último recurso para tipos sem extração suportada; formatos Office legados (`.doc`/`.ppt`/`.xls`) devolvem mensagem a pedir o formato moderno. O texto extraído é tratado como **não-confiável** (prompt injection), tal como os corpos de email. *(`.xlsx` ainda não suportado — exigiria dependência adicional.)*
- **Calendário:** consultar eventos, verificar disponibilidade, criar, editar, cancelar, responder a convites. *Comportamentos de mensagem (server-side, antes de emitir token):* (a) ao **recusar** um convite (`decline`), pergunta se quer enviar mensagem ao organizador e qual — opções: com mensagem, sem mensagem mas notifica, ou sem notificar; (b) ao **cancelar** um evento de que se é organizador, pergunta se quer mensagem própria, uma **sugestão** (que o assistente propõe e o utilizador tem de **aceitar** antes), ou nenhuma — o assistente nunca cancela com uma sugestão não aprovada. A consulta de eventos expõe o `responseStatus` do próprio (permite "quais por aceitar?"). O fuso usado é o do mailbox (`MailboxSettings.Read`), com degradação graciosa para UTC se o scope faltar. Ver [estado-user-stories §Fase 2](fase-2/estado-user-stories.md).
- **Teams (chats 1:1 e grupo):** listar chats, ler mensagens, enviar/responder.
- **Ficheiros (OneDrive/SharePoint):** pesquisar, listar, ler, upload, mover/renomear/eliminar.

> A orquestração de pedidos complexos e interligados (decomposição, dependências entre passos, ordenação, gestão de IDs, falhas parciais) está detalhada no [playbook do assistente](../src/prompts/assistant-playbook.md).

## 7. RGPD / Compliance

- **Soberania de dados:** a VPS na UE protege tokens e logs em repouso, **mas o conteúdo** (corpos de email, mensagens Teams, ficheiros) é, por construção, enviado ao modelo Claude — infra Anthropic, potencialmente fora da UE. Requisitos: **DPA com a Anthropic**, base legal de transferência internacional (SCCs/adequação), **DPIA**, e definição do papel jurídico (controlador vs subcontratante). *Decisão jurídica do cliente.*
- **Logs de auditoria:** retenção **12 meses** com purga automática; metadados apenas; acesso restrito; cifrados em repouso; pseudonimização onde viável. (Tensão reconhecida: metadados favorecem privacidade mas limitam forense.)

## 8. Testes e observabilidade

- Tenant/contas de teste dedicados; testes de integração contra o Graph.
- Health checks e métricas (latência Graph, taxa de 429, falhas de refresh).
- **Alerta de falha de refresh em massa** (uma política CA nova ou revogação pode derrubar todos os utilizadores de uma vez).

## 9. Riscos e mitigações

| Risco | Impacto | Mitigação |
|---|---|---|
| Conditional Access (dispositivo gerido) bloqueia refresh | **Alto** | PoC Fase 0; exceção de CA com escopo mínimo |
| Prompt injection via conteúdo lido | Alto | Conteúdo não-confiável; escrita só por intenção direta; sanitização HTML; risco residual aceite |
| Duplicação de envios | Médio-Alto | Idempotency keys (token de confirmação) |
| Comprometimento de tokens on-VPS | Alto | Isolamento de processo, permissões restritas; migrar a KMS depois |
| Soberania do conteúdo (Anthropic) | Médio-Alto | DPA + DPIA + base de transferência |
| Throttling Graph | Médio | Backoff + monitorização 429 |
| Fadiga de aprovação | Médio | Fricção proporcional ao risco |
| Conectividade do connector no Claude Team Mobile | Médio | PoC Fase 0 |

## 10. Faseamento

1. **Fase 0 — PoC técnico (pré-requisito):** OAuth remoto + Claude Team Desktop+Mobile + 1 utilizador + 1 tool read-only + comportamento do refresh sob CA.
2. **Fase 1 — Arquitetura de identidade + Email** (incl. anexos, two-phase, auditoria).
3. **Fase 2 — Calendário.**
4. **Fase 3 — Teams (chats).**
5. **Fase 4 — Ficheiros + multi-conta.**
6. **Futuro — Tarefas (To Do/Planner), canais Teams, dashboard de auditoria, webhooks/subscrições Graph.**

> **Fase adicional — Aprendizagem de comportamento de email (US-L.x):** módulo ortogonal que aprende, a partir das ações que o utilizador confirma (só metadados, opt-in), e **sugere** a ação habitual para emails parecidos. A execução de qualquer sugestão passa SEMPRE pelo two-phase approval existente (§3) — nunca há auto-execução. Detalhe em [docs/fase-aprendizagem/analise-funcional-aprendizagem.md](fase-aprendizagem/analise-funcional-aprendizagem.md).

## 11. Decisões em aberto (cliente)

- Conditional Access: exceção de CA vs reautenticação periódica (§2.4).
- DPA com Anthropic + DPIA + papel jurídico (§7).
- Política de CA exata do tenant (export do admin Entra) para fechar §2.4.
