# Plano de Implementação — PoC Fase 0

**Projeto:** mw-mcp-office365 (Mobiweb)
**Documento:** Plano de implementação técnica da PoC Fase 0
**Estado:** Contrato de implementação para o developer. Aprovação do coordenador pendente.
**Referências:** [Análise Funcional v1.1 §2.4 e §10](../analise-funcional-v1.1.md) · [Playbook do assistente](../../src/prompts/assistant-playbook.md)

> **Âmbito desta entrega.** APENAS a PoC Fase 0. NÃO se implementa Email/Calendário/Teams/Ficheiros. A PoC é o **esqueleto mínimo** que prova o dual-plane OAuth e permite validar, no tenant real do cliente, o bloqueador de Conditional Access (CA). A validação no tenant/VPS reais é **manual, feita pelo cliente**. A equipa entrega código + testes com Graph/Entra **mockados** + um runbook de validação manual.

---

## 1. Objetivo e critério de sucesso

### 1.1 O que a PoC tem de provar

A PoC Fase 0 é um *spike* de arquitetura: existe para **invalidar o projeto cedo e barato** se o bloqueador de CA for fatal, e para fixar o esqueleto de identidade sobre o qual as fases 1-4 vão assentar. Concretamente, deve provar:

1. **Plano A (Claude ↔ MCP).** Um servidor MCP remoto, sobre HTTP (streamable HTTP), protegido como OAuth Resource Server, descoberto e ligado por um cliente Claude remoto (Team Desktop **e** Mobile) através de:
   - Protected Resource Metadata — RFC 9728 (`/.well-known/oauth-protected-resource`).
   - Authorization Server Metadata — RFC 8414 (`/.well-known/oauth-authorization-server`).
   - Dynamic Client Registration — RFC 7591 (`POST /register`).
2. **Plano B (MCP ↔ Entra ID).** O servidor, como client confidencial, obtém tokens Graph **delegados** por Authorization Code + PKCE via `msal`, e **renova-os por refresh token**.
3. **Mapeamento de identidade.** O token do Plano A resolve, server-side, para a sessão de tokens Graph (Plano B) do utilizador correto, com isolamento estrito e modelo multi-conta-ready.
4. **Token store cifrado** em SQLite, com a camada de cifra abstraída (migrável a KMS).
5. **1 tool read-only end-to-end** (`whoami`, `User.Read`) que percorre o caminho completo: token Plano A → sessão → token Graph → chamada Graph → resposta ao Claude.
6. **Re-autenticação graciosa** quando o refresh devolve `invalid_grant`.
7. **Observabilidade mínima:** health check + log estruturado de falhas de refresh (o sinal-chave da CA).

### 1.2 Critério de sucesso (go) — verificável no tenant real pelo cliente

A PoC é **bem-sucedida** se, no tenant e VPS reais:

- **G1.** O connector remoto liga no Claude **Team Desktop** *e* **Mobile** (descoberta de metadata + DCR + Authorization Code + PKCE completam sem intervenção manual nos clientes).
- **G2.** Após o login interativo inicial (browser, dispositivo gerido), a tool `whoami` devolve a identidade do utilizador via Graph `User.Read` — o **dual-plane fecha end-to-end**.
- **G3.** Existe **pelo menos um caminho viável** para o refresh silencioso sob a CA do tenant: ou o refresh passa diretamente, ou passa após a exceção de CA de escopo mínimo (named location / identidade do servidor) que o admin Entra configura. O comportamento observado é **reproduzível e explicável** pelos logs estruturados.

### 1.3 Critério de invalidação (no-go) — o que mata ou redesenha o projeto

- **N1.** O connector remoto **não liga** no Claude Team (esp. Mobile) por limitação da plataforma — não há OAuth remoto funcional → o produto, como desenhado, não é entregável nesses clientes.
- **N2.** O refresh silencioso é **bloqueado de forma dura** pela CA (token protection / device binding) e **nenhuma exceção de escopo aceitável** o resolve → "login uma vez" é impossível; o cliente tem de aceitar reautenticação interativa periódica como constraint de design (degradação) ou reavaliar o projeto.

> A PoC **não decide** a política de CA — isso é decisão do cliente + admin Entra (v1.1 §11). A PoC **gera a evidência** (logs, comportamento observado) para essa decisão de go/no-go.

---

## 2. Arquitetura técnica concreta

### 2.1 Estrutura de pastas

Código do servidor sob `src/mcp_o365/` (package Python instalável). `src/prompts/` mantém-se intocado.

```
src/
├── mcp_o365/
│   ├── __init__.py
│   ├── config.py                 # carrega/valida configuração e segredos (.env → objeto tipado)
│   ├── app.py                    # composição: instancia ASGI app, monta rotas OAuth + MCP, DI
│   ├── server.py                 # definição do servidor MCP e registo das tools
│   ├── logging_setup.py          # configuração de logging estruturado (JSON), correlação por request
│   │
│   ├── auth/
│   │   ├── __init__.py
│   │   ├── plane_a.py            # Plano A: MCP como OAuth Resource/Auth Server p/ o Claude
│   │   ├── metadata.py           # documentos RFC 9728 + RFC 8414 (.well-known)
│   │   ├── dcr.py                # Dynamic Client Registration (RFC 7591)
│   │   ├── plane_b.py            # Plano B: client confidencial Entra ID via msal (authcode+PKCE+refresh)
│   │   └── errors.py             # exceções de auth tipadas (ReauthRequired, InvalidGrant, ...)
│   │
│   ├── identity/
│   │   ├── __init__.py
│   │   ├── mapping.py            # token Plano A → subject → session_id → sessão Graph (Plano B)
│   │   └── models.py             # dataclasses: McpPrincipal, GraphSession, LinkedAccount
│   │
│   ├── storage/
│   │   ├── __init__.py
│   │   ├── token_store.py        # CRUD de sessões/tokens em SQLite (isolado por utilizador, multi-conta)
│   │   ├── crypto.py             # interface Cipher + impl AES-GCM local (migrável a KMS)
│   │   └── schema.sql            # DDL das tabelas (sessions, linked_accounts, oauth_clients)
│   │
│   ├── graph/
│   │   ├── __init__.py
│   │   └── client.py             # wrapper HTTP fino sobre Graph (apenas /me p/ a PoC)
│   │
│   ├── tools/
│   │   ├── __init__.py
│   │   └── whoami.py             # única tool da PoC: read-only, exerce o caminho completo
│   │
│   └── observability/
│       ├── __init__.py
│       └── health.py            # endpoint /healthz + readiness
│
└── prompts/                      # (inalterado)
    └── assistant-playbook.md

tests/
├── unit/
│   ├── test_crypto.py
│   ├── test_token_store.py
│   ├── test_identity_mapping.py
│   ├── test_metadata.py
│   ├── test_dcr.py
│   └── test_plane_b_refresh.py
├── integration/
│   ├── test_oauth_flow_planeA.py # fluxo OAuth Plano A simulado (cliente fake)
│   ├── test_whoami_e2e.py        # end-to-end com Graph + Entra mockados
│   └── conftest.py               # fixtures: app de teste, mocks Graph/Entra, store temporário
└── fixtures/
    └── entra_responses/          # respostas JSON canónicas de token/erro (incl. invalid_grant)

docs/poc-fase-0/
├── plano-implementacao.md        # este documento
└── runbook-validacao-manual.md   # esqueleto na §6; QA detalha depois

pyproject.toml                    # metadados do package + dependências + config de tooling
.env.example                      # template de variáveis (sem valores) — versionado
README (atualizar src/README.md a apontar para este plano)
```

### 2.2 Responsabilidade de cada módulo

| Módulo | Responsabilidade | Notas |
|---|---|---|
| `config.py` | Ler `.env`, validar presença/forma das variáveis (§7), expor um objeto de configuração imutável. Falha rápido no arranque se faltar segredo crítico. | Pydantic Settings. Nunca loga segredos. |
| `app.py` | Composição da aplicação ASGI: instancia store, cipher, plane_a, plane_b, mapping; injeta dependências; monta rotas OAuth/metadata + transporte MCP + health. | Único ponto onde os componentes se ligam (composition root). |
| `server.py` | Define o servidor MCP, regista a tool `whoami`, liga o middleware de autenticação do Plano A (extrai e valida o token MCP de cada request). | Usa o MCP SDK oficial. |
| `logging_setup.py` | Logging estruturado JSON; `request_id`/`session_id` (pseudonimizado) por linha; **evento dedicado `refresh_failure`**. | Sem PII no log (só metadados — RGPD v1.1 §7). |
| `auth/plane_a.py` | MCP como Resource Server: valida o access token MCP recebido do Claude; resolve o `subject`. Como Auth Server: endpoints `/authorize` e `/token` que iniciam/concluem o login (delegando o login real ao Plano B / Entra). | É a "fachada" OAuth que o Claude vê. |
| `auth/metadata.py` | Serve RFC 9728 (`/.well-known/oauth-protected-resource`) e RFC 8414 (`/.well-known/oauth-authorization-server`) com os campos exigidos pelos clientes Claude. | Documentos derivados de `config.py` (issuer, endpoints, scopes). |
| `auth/dcr.py` | RFC 7591: aceita registo dinâmico de clientes Claude (`POST /register`), persiste o `client_id` emitido em `oauth_clients`. | Política de registo mínima mas funcional. |
| `auth/plane_b.py` | Client confidencial Entra via `msal.ConfidentialClientApplication`: gera URL de autorização (authcode + PKCE), troca o código por tokens Graph, **faz refresh**, normaliza erros (deteta `invalid_grant`). | Núcleo do bloqueador de CA: o refresh vive aqui. |
| `auth/errors.py` | Exceções tipadas: `ReauthRequired`, `InvalidGrant`, `ConsentRequired`, `UpstreamAuthError`. | Permite à tool/serviço reagir graciosamente. |
| `identity/mapping.py` | O elo crítico: dado o `subject` do token Plano A, devolve a `GraphSession` (Plano B) certa. Cria/atualiza sessões; aplica isolamento por utilizador; suporta N contas por principal. | `subject (Plano A) → session → linked_account(s) → tokens Graph`. |
| `identity/models.py` | `McpPrincipal` (quem o Claude diz que é), `GraphSession`, `LinkedAccount` (uma conta O365 ligada, com `tenant_id`/`home_account_id`/scopes). | Multi-conta desde o início: `LinkedAccount` é 1:N por sessão. |
| `storage/token_store.py` | Persistência SQLite: criar/ler/atualizar/expirar sessões e contas ligadas; guardar tokens **cifrados**; queries sempre filtradas por `subject`/`account_id` (sem acesso cruzado). | Toda a escrita de token passa pelo `Cipher`. |
| `storage/crypto.py` | Interface `Cipher` (`encrypt`/`decrypt`) + impl `LocalAesGcmCipher` com chave de `config`. Ponto único de cifra para trocar por KMS depois. | Abstração obrigatória (v1.1 §4). |
| `storage/schema.sql` | DDL: `sessions`, `linked_accounts`, `oauth_clients`. Índices por `subject`. | Migração simples (aplicada no arranque se as tabelas não existirem). |
| `graph/client.py` | Wrapper fino sobre Microsoft Graph; na PoC só `GET /me`. Recebe um access token, devolve dados; trata 401/403 e 429 (`Retry-After`) de forma mínima. | Não conhece tokens nem store — recebe o token pronto. |
| `tools/whoami.py` | A única tool MCP: pega no `subject` do request → `mapping` → token Graph (refresh se necessário via `plane_b`) → `graph.client.me()` → devolve `displayName`/`userPrincipalName`/`id`. Em `InvalidGrant`, devolve mensagem de re-login. | Prova o dual-plane completo. Read-only, sem two-phase. |
| `observability/health.py` | `GET /healthz` (liveness) e readiness (DB acessível, config carregada). Não toca em Entra/Graph (não consome rate limit). | |

### 2.3 Fluxos-chave (resumo)

**Ligação inicial (Plano A + B):**
```
Claude → GET /.well-known/oauth-protected-resource         (RFC 9728)
Claude → GET /.well-known/oauth-authorization-server       (RFC 8414)
Claude → POST /register                                    (RFC 7591) → client_id
Claude → GET /authorize  → MCP gera authcode+PKCE p/ Entra → redirect ao Entra (browser do user)
   [user faz login no dispositivo GERIDO → CA passa]
Entra → redirect → MCP /callback → plane_b troca code por tokens Graph
   → mapping cria session + linked_account → token_store guarda (cifrado)
   → MCP emite access token do Plano A para o Claude
```

**Chamada à tool (caminho que valida tudo):**
```
Claude → MCP (tool whoami, com access token Plano A)
   server.py valida token (plane_a) → subject
   mapping(subject) → GraphSession → access token Graph
      se expirado → plane_b.refresh()  ← AQUI o refresh enfrenta a CA
         se invalid_grant → log refresh_failure + ReauthRequired → Claude pede re-login
   graph.client.me(access_token) → { displayName, upn, id }
   → resposta ao Claude
```

---

## 3. Escolha justificada de bibliotecas

| Necessidade | Package | Porquê (1 linha) |
|---|---|---|
| MCP SDK Python | **`mcp`** (Model Context Protocol Python SDK oficial, com FastMCP) | SDK oficial, suporta transporte **streamable HTTP** e o padrão de **OAuth protected resource** exigidos pela v1.1. |
| Fluxo Entra ID (Plano B) | **`msal`** (Microsoft Authentication Library) | Biblioteca oficial Microsoft para Authorization Code + PKCE e gestão de refresh tokens — decisão do cliente. |
| Framework HTTP / ASGI | **`starlette`** (servida por **`uvicorn`**) | Leve, é a base ASGI sobre que o MCP SDK monta o transporte HTTP; permite anexar facilmente as rotas `.well-known`/OAuth/health. |
| Cliente HTTP p/ Graph e Entra | **`httpx`** | Cliente moderno (sync+async), timeouts e retries explícitos; usado pelo `graph/client.py`. |
| Cifra de tokens | **`cryptography`** (AES-256-GCM via `AESGCM`) | Primitiva AEAD auditada e padrão; encapsulada atrás da interface `Cipher` para migrar a KMS sem tocar no resto. |
| SQLite | **`sqlite3`** (stdlib) | Driver embutido, zero dependências externas; suficiente para o token store cifrado da PoC. |
| Configuração / validação | **`pydantic-settings`** | Carrega `.env` com validação tipada e falha-rápido se faltar segredo crítico. |
| Testes | **`pytest`** + **`respx`** (mocks httpx) + **`pytest-asyncio`** | `pytest` padrão; `respx` intercepta chamadas Graph/Entra para testar sem rede real; async para o transporte. |

> O nome exato do package do MCP SDK a fixar no `pyproject.toml` deve ser confirmado pelo developer na primeira tarefa (T0) contra o índice oficial; o plano assume `mcp`.

---

## 4. Decomposição em tarefas (para o developer)

> Sequência recomendada. Dependências marcadas em **Dep.** Cada tarefa tem uma **Definição de Pronto (DoD)** verificável. Toda a entrega usa Graph/Entra **mockados**; nenhuma tarefa requer o tenant real.

| # | Tarefa | Dep. | Definição de Pronto (DoD) |
|---|---|---|---|
| **T0** | **Scaffold do projeto.** Criar `pyproject.toml`, package `src/mcp_o365/`, `tests/`, `.env.example`, configurar `pytest`, lint/format, e **confirmar o nome do package do MCP SDK**. | — | `pip install -e .` instala; `pytest` corre (0 testes ou 1 trivial passa); `.env.example` lista todas as variáveis da §7 sem valores; lint passa. |
| **T1** | **`config.py`.** Carregar e validar configuração via pydantic-settings; falha-rápido com mensagem clara se faltar segredo. | T0 | Teste unitário: config válida carrega; config a faltar `ENTRA_CLIENT_SECRET` levanta erro no arranque. Segredos nunca aparecem no `repr`/log. |
| **T2** | **`logging_setup.py`.** Logging estruturado JSON com `request_id`; evento dedicado `refresh_failure` com campos fixos (`event`, `subject_hash`, `account_id`, `reason`). Sem PII. | T0 | Teste: emitir `refresh_failure` produz JSON com os campos esperados; nenhum campo contém token/UPN em claro. |
| **T3** | **`storage/crypto.py`.** Interface `Cipher` + `LocalAesGcmCipher` (AES-256-GCM). | T1 | Teste: `decrypt(encrypt(x)) == x`; nonce único por operação; texto cifrado difere a cada chamada; chave errada falha a decifrar. |
| **T4** | **`storage/schema.sql` + `token_store.py`.** Tabelas `sessions`/`linked_accounts`/`oauth_clients`; CRUD com tokens cifrados; **todas as queries filtram por `subject`/`account_id`**; aplicar schema no arranque. | T3 | Testes: criar sessão; guardar/ler token (cifrado em repouso — verificar que a coluna não contém o token em claro); **isolamento**: utilizador A não lê sessão de B; suporte multi-conta (2 `linked_accounts` na mesma sessão). |
| **T5** | **`identity/models.py` + `identity/mapping.py`.** Resolver `subject` (Plano A) → `GraphSession` (Plano B); criar/atualizar; marcar expirada; selecionar conta (default + multi). | T4 | Testes: mapping devolve a sessão certa por subject; subject desconhecido → `None`/`ReauthRequired`; seleção entre 2 contas funciona; sessão marcada expirada não é devolvida como ativa. |
| **T6** | **`auth/plane_b.py` + `auth/errors.py`.** `msal.ConfidentialClientApplication`: gerar URL de autorização (authcode+PKCE), trocar código por tokens, **refresh**, normalizar `invalid_grant` → `InvalidGrant`/`ReauthRequired`. Entra **mockado** (`respx` + fixtures). | T5 | Testes: troca de code devolve tokens e cria sessão; refresh com token válido renova; **resposta `invalid_grant` → `ReauthRequired` + log `refresh_failure`**; `consent_required` → `ConsentRequired`. |
| **T7** | **`auth/metadata.py`.** Servir RFC 9728 e RFC 8414 com os campos derivados de `config`. | T1 | Testes: `GET /.well-known/oauth-protected-resource` e `.../oauth-authorization-server` devolvem JSON com os campos obrigatórios (issuer, endpoints, scopes, `registration_endpoint`). |
| **T8** | **`auth/dcr.py`.** RFC 7591: `POST /register` emite e persiste `client_id`. | T4, T7 | Teste: registo devolve `client_id` válido e persiste em `oauth_clients`; pedido malformado → 400. |
| **T9** | **`auth/plane_a.py`.** Resource Server (validar token MCP → `subject`) + Auth Server (`/authorize`, `/token`, `/callback`) que orquestram o Plano B. | T6, T8 | Teste de integração (cliente OAuth **fake**): fluxo `authorize → callback → token` completa e produz um access token MCP resolúvel para um `subject`; request sem token válido → 401 com `WWW-Authenticate` apontando para a metadata. |
| **T10** | **`graph/client.py`.** Wrapper `me()` sobre `GET /me`; trata 401/403 e 429 (`Retry-After`). Graph **mockado**. | T0 | Testes: `me()` devolve campos esperados de uma resposta mockada; 401 → erro tipado; 429 respeita `Retry-After` (sem dormir de verdade no teste). |
| **T11** | **`tools/whoami.py` + `server.py`.** Registar a tool no servidor MCP; ligar middleware de auth do Plano A; caminho completo subject→mapping→(refresh?)→graph. | T9, T10, T5 | Teste E2E (Entra+Graph mockados): chamada à tool com sessão válida devolve identidade; com refresh necessário, renova e devolve; com `invalid_grant`, devolve mensagem de re-login (não rebenta). |
| **T12** | **`observability/health.py` + `app.py`.** `/healthz` + readiness; `app.py` compõe tudo (store, cipher, planes, mapping, MCP, rotas, health). | T11 | Teste: `GET /healthz` → 200; readiness falha se a DB estiver inacessível; app arranca com `.env.example` preenchido com valores fake. |
| **T13** | **Runbook de validação manual.** Preencher `runbook-validacao-manual.md` a partir do esqueleto (§6), com os passos concretos e os pontos de observação dos logs. | T12 | Documento revisto pelo QA: passos numerados, pré-requisitos, e critérios G1-G3/N1-N2 mapeados a observações concretas. |

> **Caminho crítico:** T0 → T1 → T3 → T4 → T5 → T6 → T9 → T11 → T12. T2, T7, T8, T10 são paralelizáveis após as suas dependências.

---

## 5. Estratégia de teste (para o QA)

### 5.1 O que é testável automaticamente nesta entrega

Com Graph/Entra **mockados** (`respx` + fixtures JSON), tudo isto fecha em CI sem tenant real:

- **Unit:**
  - `crypto`: round-trip, unicidade de nonce, falha com chave errada.
  - `token_store`: cifra em repouso, **isolamento por utilizador**, multi-conta, expiração de sessão.
  - `identity/mapping`: resolução por subject, subject desconhecido, seleção de conta.
  - `config`: falha-rápido sem segredos; ausência de segredos em logs/`repr`.
  - `logging`: evento `refresh_failure` com schema correto e sem PII.
- **Integração (mockada):**
  - **Fluxo OAuth Plano A simulado:** um cliente OAuth *fake* exerce metadata (RFC 9728/8414) → DCR (RFC 7591) → `authorize/callback/token`, e obtém um token MCP utilizável.
  - **Plano B:** troca de code, **refresh bem-sucedido**, e **`invalid_grant` → reauth graciosa + log** (com fixtures de resposta Entra canónicas, incl. o JSON de erro `invalid_grant`).
  - **whoami E2E:** caminho completo com sessão válida, com refresh-no-meio, e com refresh falhado.
  - **health:** liveness/readiness.

### 5.2 O que NÃO é testável aqui — só validável manualmente no tenant real

São exatamente os pontos que justificam a PoC. **Não tentar automatizar**; são manuais e da responsabilidade do cliente (runbook, §6):

- **Comportamento real do connector no Claude Team Desktop e Mobile** (G1/N1) — depende da plataforma Claude e não é mockável.
- **Comportamento real do refresh silencioso sob a política de Conditional Access do tenant** (G3/N2) — depende do Entra real, do dispositivo gerido e da política de CA. Os mocks provam que o *código* trata refresh e `invalid_grant`; **não** provam o que o Entra real faz sob CA.
- **Admin consent global** efetivo no tenant.
- **TLS / Cloudflare Origin Certificate (Full strict)** e o redirect URI no domínio final.

### 5.3 Princípios de QA

- **Sem segredos reais em CI.** Fixtures e `.env` de teste com valores fake.
- **Testes determinísticos:** sem rede real, sem `sleep` real (429 testado com relógio injetável/mock).
- **Cobertura prioritária:** crypto, isolamento do token_store e o caminho `invalid_grant` (é o que prova a reauth graciosa, central ao risco de CA).
- **Marcar claramente** os testes de integração que simulam OAuth para não serem confundidos com validação real.

---

## 6. Esqueleto do runbook de validação manual

> Vai para `docs/poc-fase-0/runbook-validacao-manual.md`. O QA detalha cada passo; aqui fica a espinha. Executado **pelo cliente** no tenant + VPS reais.

**Pré-requisitos**
- Acesso admin ao tenant Entra ID; export da política de CA atual.
- VPS com o domínio `mw-mcp-office365.mobiweb.pt` atrás de Cloudflare (Origin Certificate, Full strict).
- 1 utilizador de teste num **dispositivo gerido/compliant**.
- Claude **Team** com Desktop e Mobile disponíveis.

**Passos**
1. **Registar a app no Entra ID** — client confidencial; redirect URI = `https://mw-mcp-office365.mobiweb.pt/callback` (domínio final); gerar `client secret`; configurar scopes delegados (`User.Read`, `offline_access`, `openid`, `profile`).
2. **Admin consent global** dos scopes da PoC (tenant-wide).
3. **Configurar segredos na VPS** (§7) e arrancar o servidor; confirmar `GET /healthz` → 200.
4. **Validar a metadata pública:** `GET /.well-known/oauth-protected-resource` e `.../oauth-authorization-server` respondem com TLS válido pelo domínio final.
5. **Ligar o connector no Claude Team Desktop** — observar descoberta de metadata + DCR + login interativo (browser, dispositivo gerido) → executar a tool `whoami` → confirma identidade. **(G1, G2)**
6. **Ligar o connector no Claude Team Mobile** — repetir o login e `whoami`. **(G1 — ponto de maior risco de plataforma)**
7. **Observar o refresh sob CA:** forçar a expiração do access token (esperar / reduzir validade no Entra) e voltar a chamar `whoami`. Observar nos **logs estruturados** se o `plane_b.refresh()` **passa** ou regista **`refresh_failure`**. **(G3 / N2)**
8. **Se o refresh for bloqueado:** o admin Entra aplica a **exceção de CA de escopo mínimo** (named location = IP da VPS, ou identidade do servidor) e repete o passo 7. Registar o antes/depois.
9. **Testar a reauth graciosa:** revogar a sessão no Entra (ou aguardar expiração do refresh) → confirmar que o Claude recebe o pedido de **re-login** em vez de uma falha silenciosa.
10. **Registar o veredito go/no-go** mapeando observações a G1-G3 / N1-N2 (§1.2/1.3), e anexar os logs relevantes.

---

## 7. Configuração e segredos

Tudo via `.env` (já no `.gitignore`). `.env.example` versionado com as chaves e **sem valores**.

**Vão para `.env` (segredos / específicos do ambiente):**

| Variável | Descrição |
|---|---|
| `ENTRA_TENANT_ID` | ID do tenant Entra. |
| `ENTRA_CLIENT_ID` | App (client) ID do registo confidencial. |
| `ENTRA_CLIENT_SECRET` | **Segredo** do client confidencial. |
| `ENTRA_AUTHORITY` | Authority Entra (ex.: `https://login.microsoftonline.com/<tenant>`). |
| `OAUTH_REDIRECT_URI` | Redirect do callback no domínio final (`https://.../callback`). |
| `GRAPH_SCOPES` | Scopes delegados da PoC (`User.Read offline_access openid profile`). |
| `MCP_ISSUER_URL` | Issuer público do servidor (base das metadata, domínio final). |
| `MCP_PUBLIC_BASE_URL` | URL base público do servidor MCP. |
| `TOKEN_STORE_PATH` | Caminho do ficheiro SQLite (fora da raiz web). |
| `TOKEN_ENCRYPTION_KEY` | **Chave** de cifra AES-256 (base64). Migrável a KMS depois. |
| `LOG_LEVEL` | Nível de log (`INFO`/`DEBUG`). |
| `BIND_HOST` / `BIND_PORT` | Host/porta de escuta do uvicorn. |

**Não são segredos mas são config de ambiente:** `MCP_ISSUER_URL`, `MCP_PUBLIC_BASE_URL`, `OAUTH_REDIRECT_URI`, `BIND_*`, `LOG_LEVEL` — podem estar no `.env` por conveniência.

**Regras:** segredos nunca em logs nem em mensagens de erro; chave de cifra e ficheiro SQLite **fora da raiz web** e com permissões restritas; nenhum `.pem`/`tokens.db`/`.env` commitado (já coberto pelo `.gitignore`).

---

## 8. Riscos abertos da PoC e o que significam para o go/no-go

| Risco | Descrição | Sinal observável | Implicação go/no-go |
|---|---|---|---|
| **R1 — Refresh bloqueado pela CA** | A CA de dispositivo gerido rejeita o refresh silencioso do servidor (não-gerido), possivelmente com token protection/binding. | Logs `refresh_failure` com `invalid_grant`/erro de CA no passo 7 do runbook, **mesmo após** a exceção de escopo mínimo. | **Potencial no-go (N2)** ou degradação a reauth periódica. É o risco central da PoC. |
| **R2 — Connector não liga no Claude Team (Mobile)** | A plataforma Claude pode não suportar o fluxo OAuth remoto + DCR em todos os clientes, especialmente Mobile. | Falha de ligação/descoberta no passo 6. | **No-go parcial (N1)** para esse cliente. |
| **R3 — Exceção de CA inaceitável** | A única exceção que faz o refresh passar é demasiado ampla para o cliente aceitar (ex.: alarga superfície de ataque). | Refresh só passa com exceção que o cliente recusa por política de segurança. | Decisão do cliente: aceitar reauth periódica ou não avançar. |
| **R4 — DCR / metadata incompatíveis** | Os clientes Claude esperam variantes específicas de RFC 9728/8414/7591 que divergem da nossa implementação. | Cliente Claude rejeita a metadata/registo no passo 4-5. | Ajuste de implementação (não fatal), mas pode atrasar; validar cedo. |
| **R5 — Admin consent insuficiente** | Consent global não cobre os scopes ou requer passo extra por utilizador. | Ecrã de consentimento por utilizador no login (passo 5). | Ajuste de configuração no Entra; não fatal. |
| **R6 — Nome/versão do MCP SDK** | O package/versão assumido (`mcp`) pode divergir do disponível, afetando o transporte HTTP/OAuth. | Falha em T0. | Resolvido em T0 (fixar a dependência antes de tudo). |

> **Nota de leitura:** R1 e R2 são os únicos riscos com potencial de **no-go real**; os restantes são ajustáveis. É por eles que a PoC existe — todo o código é instrumental a gerar evidência limpa sobre R1 e R2.

---

*Fim do plano. O developer implementa à risca; o QA usa a §5 para o plano de testes e detalha o runbook (§6). Alterações de âmbito carecem de aprovação do coordenador.*
