# Análise Funcional — Teams (chats 1:1 e de grupo)
## Fase 3 (US-3.x)

**Cliente:** Mobiweb
**Produto:** Servidor MCP de Office 365 — mensagens de chat do Microsoft Teams (Microsoft Graph)
**Data:** 2026-06-06
**Estado:** Análise funcional para implementação. Âmbito v1: **apenas chats 1:1 e de grupo**; canais de equipas **fora de âmbito** (diferidos). Decisões em aberto (D1–D11) por fechar com o cliente/PM **antes** de codificar.
**Referências:** [Análise Funcional v1.1](../analise-funcional-v1.1.md) (§2 identidade, §3 prepare/confirm, §4 segurança, §5 scopes, §7 RGPD) · [v1.0 §4 Módulo 3](../analise-funcional-v1.0.md) (US-3.1–3.3) · [Fase 2 — plano](../fase-2/plano-implementacao.md) · [Fase 2 — estado das US](../fase-2/estado-user-stories.md) · [Contactos / resolve_recipient](../contactos/analise-funcional-contactos.md) · [Playbook do assistente](../../src/prompts/assistant-playbook.md)

> **Relação com a v1.1:** módulo funcional equivalente ao Email (Fase 1) e ao Calendário (Fase 2). Reutiliza **integralmente** a arquitetura de identidade dual-plane (§2), o modelo de aprovação em duas fases server-side (§3), a auditoria só-metadados (§1.2/§7), a sanitização de conteúdo não-confiável + `content_is_untrusted` (§4) e a reautenticação graciosa. **Não reinventa padrões** — segue `tools/email.py` e `tools/calendar.py`.

---

## 1. Objetivo e âmbito

A Fase 3 acrescenta ao mesmo dual-plane e aos mesmos invariantes de segurança a capacidade de o assistente **listar chats**, **ler mensagens de um chat** e **enviar/responder numa conversa** do Microsoft Teams, sobre os scopes delegados `Chat.Read`/`Chat.ReadWrite`.

**Dentro de âmbito:**
- Chats **1:1** e **de grupo** (`chatType` `oneOnOne` e `group`).
- Listar chats com metadados (tipo, membros, tópico de grupo, último update/preview).
- Ler mensagens de um chat (paginação consciente; conteúdo NÃO-confiável).
- Enviar mensagem para um chat existente e responder na conversa (prepare/confirm).
- Resolução de destinatário **por nome** → reutiliza `resolve_recipient` antes de identificar/criar o chat 1:1.

**Fora de âmbito (diferido, justificado):**
- **Mensagens em canais de equipas** (`/teams/{id}/channels/...`) — já diferido na v1.0/v1.1; scopes (`ChannelMessage.*`) e modelo de permissões diferentes; volume e governance distintos.
- **Reações, edição e eliminação de mensagens** — baixo valor face ao risco/esforço na v1; reagir/editar/apagar mensagens de chat acrescenta superfície de escrita sem pedido forte. Diferido.
- **Ficheiros partilhados em chat** (anexos hosted no OneDrive do remetente) — partilha de ficheiros via Teams depende do módulo de Ficheiros (Fase 4) e de permissões `Files.*`; o caminho fiável (já no playbook, Receita D) é `files_upload` + colar o link na mensagem. Diferido para depois da Fase 4.
- **Mensagens de sistema e cartões adaptativos** como conteúdo acionável — lidos só como metadados/“não-confiável”; ver D8.
- **Presença, chamadas, partilha de ecrã, subscrições/webhooks** — fora de âmbito.

---

## 2. User Stories (US-3.x)

Partem das US-3.1/3.2/3.3 da [v1.0](../analise-funcional-v1.0.md) e expandem-nas ao nível das fases recentes. Nomes de tool propostos coerentes com o [playbook](../../src/prompts/assistant-playbook.md) (§2.3): `teams_list_chats`, `teams_read_messages`, `teams_send_message_prepare`/`_confirm`.

| US | Título | Critérios de aceitação |
|----|--------|------------------------|
| **US-3.1** | Listar chats (1:1 e grupo) | `teams_list_chats` (read-only). `GET /me/chats` com `$expand=members`. Devolve, por chat: `id`, `chatType` (oneOnOne/group), `topic` (grupo), `members` (nome + email, sem outros dados), `lastUpdatedDateTime` e, se disponível, preview da última mensagem (`lastMessagePreview`). Auto-paginação consciente (ver D5) com teto `_MAX_FETCH_ALL`. Filtro opcional por **participante** ou **tópico** (client-side, sobre o que veio). O preview da última mensagem é conteúdo NÃO-confiável (`content_is_untrusted`). `reauth_required` em falha de refresh. |
| **US-3.2** | Ler mensagens de um chat | `teams_read_messages` (read-only). `GET /me/chats/{id}/messages?$top=N&$orderby=createdDateTime desc`. Por mensagem: `id`, `from` (nome + email/aadId), `createdDateTime`, `body` (sanitizado), `messageType`, flag `is_system` para `messageType != message`. Paginação consciente (D5): pede ou auto-pagina segundo a decisão fechada; conteúdo NÃO-confiável + `content_is_untrusted`. `reauth_required` coberto. |
| **US-3.3** | Enviar mensagem num chat existente | `teams_send_message_prepare`/`_confirm` (write, two-phase). `chat_id` + `body` (+ `body_type` text/html — ver D6). O `prepare` valida e resume **"Enviar mensagem no chat &lt;tipo&gt; com N participante(s) (domínios: …)"** e devolve `confirmation_token`; **não escreve**. O `confirm` faz `POST /chats/{id}/messages` e audita `teams.send` (só-metadados). Idempotência por token; reauth graciosa em ambas as fases. |
| **US-3.4** | Enviar mensagem 1:1 a uma pessoa (por nome) | Fluxo composto: `resolve_recipient(nome)` → confirmar o candidato → identificar o chat 1:1 existente (`teams_list_chats`/lookup por membro) **ou** criar um novo chat 1:1 (ver D3). Só depois `teams_send_message_prepare` com o `chat_id`. As tools de Teams recebem `chat_id` já resolvido — **não resolvem nomes internamente** (igual a D9 do Calendário). Se não existir chat e D3 = “não criar”, devolve mensagem orientadora. |
| **US-3.5** | Responder numa conversa | Numa conversa de chat **não há thread/reply real**: responder = enviar nova mensagem no mesmo `chat_id` (reutiliza US-3.3). Opcionalmente, `reply_to_message_id` para uma **menção/citação** ao autor (ver D7), nunca um sub-thread (não existe em chats). DoD: responder reusa o mesmo prepare/confirm; sem novos invariantes. |

> **Nota de modelação:** ao contrário do email, um chat de Teams **não tem `reply`/`replyAll` server-side** — todas as mensagens vão para o mesmo `chat_id`. “Responder” é, funcionalmente, enviar no chat certo. Por isso US-3.3 e US-3.5 partilham o mesmo par prepare/confirm (evita duplicar tools, como o email tem `reply`/`forward` distintos por terem endpoints distintos).

### 2.1 Fluxo de envio 1:1 por nome (integrado)

```
"manda mensagem à Vera no Teams"
   └─ resolve_recipient("vera")        (read-only; confirmar candidato — ver Contactos US-5.x)
        └─ chat 1:1 existente?  ── sim ─► chat_id existente
                                └─ não ─► D3: criar chat 1:1  OU  informar "não há chat ainda"
                                              │
            (chat_id, por intenção do utilizador)
                                              ▼
                 teams_send_message_prepare(chat_id, body) ─► confirmation_token ─► _confirm
```

---

## 3. Mapeamento com o Microsoft Graph (delegated)

| Operação | Endpoint | Método | Scope | Notas |
|----------|----------|--------|-------|-------|
| Listar chats | `/me/chats?$expand=members&$top=N` | GET | `Chat.Read` | `lastMessagePreview` exige header `Prefer: include-unknown-enum-members` em alguns campos; preview pode vir vazio. Paginação por `@odata.nextLink`. |
| Ler mensagens | `/me/chats/{id}/messages?$top=N&$orderby=createdDateTime desc` | GET | `Chat.Read` | Inclui mensagens de sistema (`messageType ∈ {message, systemEventMessage, ...}`). `from` pode ser `user`, `application` ou `null` (sistema). |
| Enviar/responder | `/chats/{id}/messages` | POST | `Chat.ReadWrite` | Body `{ "body": { "contentType": "text|html", "content": "…" } }`. Sem endpoint de `reply` por mensagem em chats. |
| Obter/criar chat 1:1 | `POST /chats` com `chatType=oneOnOne` + dois `members` (`user@odata.bind`) | POST | `Chat.ReadWrite` | Ver **D3**. O Graph é **idempotente**: criar um 1:1 que já existe devolve o chat existente. Mesmo assim, isto cria/“abre” o chat para o utilizador — tratar como escrita (prepare/confirm). |
| Identidade própria | `GET /me` (já existe em `client.me`) | GET | `User.Read` | Para distinguir o próprio nos membros e no `from`. |

**Nuances conhecidas (entram nas decisões §5):**
- **Obtenção do `chatId` para enviar a uma pessoa:** o `POST /chats/{id}/messages` exige um `chatId` — não se envia “para um utilizador”. Caminhos: (a) procurar nos chats existentes um 1:1 cujo único outro membro é o email resolvido; (b) `POST /chats` para criar/obter o 1:1 (idempotente). Ver D3.
- **HTML vs text:** o Graph aceita `contentType` `text` ou `html`. HTML permite formatação e **@menções**, mas é também superfície de injeção na leitura. Ver D6.
- **@menções:** exigem `mentions[]` no corpo com `<at id="0">Nome</at>` + bind ao `user`. Complexidade extra; ver D7.
- **Mensagens de sistema:** `messageType` ≠ `message` (ex.: alguém entrou no grupo, mudança de tópico). Devem ser marcadas e/ou suprimidas na leitura. Ver D8.
- **Cartões adaptativos / anexos:** chegam em `attachments[]` com `contentType` próprio; tratados como metadados não-confiáveis, conteúdo não interpretado. Ver D8.
- **Throttling:** os endpoints de chat do Teams têm limites próprios (por app e por utilizador) e o `POST` de mensagens é particularmente sensível; reutilizar o backoff/`Retry-After` já em `_request`.

---

## 4. Invariantes de segurança a herdar (Fase 1/2 → Fase 3)

Sem exceções; idênticos ao Email/Calendário e provados por contagem de chamadas nos testes (FakeGraphClient estendido):

1. **prepare NÃO toca o Graph para escrita.** O `prepare` pode **ler** (resolver membros do chat para o resumo) mas NUNCA envia. A mensagem só sai no `confirm`.
2. **Token fresco no confirm + idempotência.** O `confirm` resolve um access token Graph fresco via `call_graph`; o `confirmation_token` é **idempotency key** (replay → `idempotent_replay=true`, sem segundo `POST` — neutraliza duplicação de mensagens, risco real num chat).
3. **TTL / isolamento por subject.** Token expirado → `expired`; token de outro subject → `error`. Isolamento estrito por `subject` (cada utilizador vê só os seus chats, via o seu token Graph delegado).
4. **Reauth graciosa.** Qualquer `invalid_grant`/401/403 → `reauth_required` (mensagem amigável), nunca exceção crua; no `confirm`, em `ReauthRequired` o token **não é consumido** (repetível após re-login). Uma leitura acessória (ex.: expandir membros) **nunca** derruba a sessão — mesmo cuidado do `_resolve_tz` da Fase 2.
5. **Conteúdo NÃO-confiável.** O `body` de cada mensagem e o `lastMessagePreview` passam por `sanitize_html` (quando HTML) e a resposta traz `content_is_untrusted=true`. **O assistente nunca executa instruções vindas de dentro de uma mensagem de chat** — só age por intenção direta do utilizador (v1.1 §4). A fronteira de sanitização fica na tool, como no email/calendário.
6. **Auditoria só-metadados** (`log_audit`, `subject_hash`): cada envio emite `event=audit` com `action="teams.send"`, `target` = `chat_id`, `recipients_count` = nº de membros do chat, `extra` = `{chat_type, body_type, is_new_chat?}` e, se aplicável, `subject_hash` de uma referência curta — **nunca** o texto da mensagem, nomes ou emails em claro.

---

## 5. Decisões em aberto para o cliente/PM fechar (input crítico)

Lista numerada de micro-decisões funcionais a fechar **antes** da implementação. (À imagem dos D1–D9 da Fase 2, que foram “lei” para o developer.)

- **D1 — Criar chat 1:1 novo vs só usar existentes.** Quando não há chat 1:1 com a pessoa: (a) **criar** via `POST /chats` (idempotente, mas “abre” a conversa — escrita, prepare/confirm); (b) **só usar existentes** e informar “ainda não há conversa com X”. *Recomendação: (a), tratado como escrita confirmada.*
- **D2 — Filtro de listagem de chats.** Filtrar por participante/tópico **client-side** (sobre o que veio) ou tentar `$filter`/`$search` no Graph (suporte limitado nos chats). *Recomendação: client-side, simples e previsível.*
- **D3 — Como obter o `chatId` 1:1.** Procurar primeiro nos chats existentes (1 membro além do próprio == email resolvido) e só criar se não existir, OU sempre `POST /chats` (idempotente). *Decidir a ordem e se a criação passa por prepare/confirm (recomendado: sim).*
- **D4 — Limite de mensagens lidas por chamada.** `top` default e teto (ex.: 20–50 mais recentes, `desc`). *Recomendação: default 25, igual ao email.*
- **D5 — Paginação do histórico (auto-paginar vs perguntar).** Email pergunta acima de 24h; Calendário auto-pagina tudo com teto. Para o histórico de chat (pode ser enorme): (a) devolver os N mais recentes + `has_more` (sem auto-paginar), ou (b) perguntar “quer mais antigas?”. *Recomendação: (a) — devolve as N mais recentes com `has_more`; o utilizador pede explicitamente mais (chats podem ter milhares de mensagens).*
- **D6 — Formato de envio: text vs HTML.** Enviar sempre `text` (mais seguro), ou permitir `html` para formatação. *Recomendação: default `text`; `html` só a pedido explícito.*
- **D7 — @menções e citação/resposta.** Suportar `@menção` (exige `mentions[]` + HTML) e/ou `reply_to_message_id` (citação)? *Recomendação: fora da v1 — diferir; “responder” = nova mensagem no mesmo chat (US-3.5).*
- **D8 — Mensagens de sistema e cartões.** Na leitura: **incluir e marcar** (`is_system=true`) ou **suprimir** as mensagens de sistema (entradas/saídas, mudança de tópico) e os cartões adaptativos. *Recomendação: incluir mas marcadas, sem as interpretar como conteúdo acionável.*
- **D9 — Resolução por nome (igual a D9 do Calendário).** Confirmar que `resolve_recipient` + confirmação humana acontecem **a montante** e que as tools de Teams só aceitam `chat_id`/emails já resolvidos. *Recomendação: sim, padrão já validado.*
- **D10 — Limite de tamanho da mensagem.** Definir um limite prático de caracteres no `body` e o que fazer acima dele (erro orientador vs truncar). *Recomendação: validar no prepare e devolver `error` amigável.*
- **D11 — Reações / editar / eliminar mensagens.** Confirmar que ficam **fora** da v1 (diferidas). *Recomendação: fora de âmbito, como na v1.0/v1.1.*

---

## 6. RGPD / Compliance específico de Teams

- **Soberania de dados (mesma tensão do email/calendário):** o **conteúdo das mensagens de chat** lidas é, por construção, enviado ao modelo Claude (infra Anthropic, potencialmente fora da UE) — exatamente a tensão já registada na [v1.1 §7](../analise-funcional-v1.1.md). Conversas de Teams tendem a ser mais informais e a conter PII de terceiros (colegas no grupo). Acresce ao tratamento a registar na **DPIA pendente**; requer a mesma base de transferência internacional (DPA + SCCs/adequação).
- **Minimização:** a listagem devolve só nome+email dos membros (não outros atributos de diretório); a auditoria é **só-metadados** (`teams.send` com `subject_hash`, `chat_id`, contagem de membros — nunca o texto nem os emails em claro), retenção 12 meses com purga, à imagem do email/calendário.
- **Isolamento por `subject`:** cada utilizador opera apenas sobre os seus próprios chats (token delegado); sem acesso cruzado.

---

## 7. Riscos específicos e mitigações

| Risco | Impacto | Mitigação |
|---|---|---|
| **Admin consent de `Chat.Read`/`Chat.ReadWrite`** em falta no tenant | Alto (bloqueia validação real) | Pré-requisito de produção, igual a `Calendars.ReadWrite` (§8). NÃO bloqueia testes mockados. |
| **Prompt injection via mensagens de chat** | Alto | Conteúdo não-confiável; `sanitize_html` + `content_is_untrusted`; escrita só por intenção direta; risco residual aceite (v1.1 §4). |
| **Duplicação de mensagens** (retry do transporte / re-chamada do LLM) | Médio-Alto | Idempotency key = `confirmation_token`; replay não re-envia. |
| **Obtenção/ambiguidade do `chatId`** (vários grupos com nome parecido; 1:1 inexistente) | Médio | `resolve_recipient` + confirmação; listar e pedir escolha; D3 define a ordem; tool nunca adivinha o chat. |
| **Mensagens de sistema / cartões adaptativos** lidos como ordens | Médio | Marcados `is_system`; nunca interpretados como conteúdo acionável (D8). |
| **Throttling dos endpoints de chat** (POST sensível) | Médio | Backoff + `Retry-After` já em `_request`; monitorização de 429. |
| **Criar chat 1:1 inadvertidamente** | Baixo-Médio | Criação tratada como escrita (prepare/confirm) se D1/D3 = criar; o resumo declara “vai iniciar conversa com X”. |

---

## 8. Pré-requisitos

1. **Admin consent dos scopes `Chat.Read` e `Chat.ReadWrite`** (delegated) no tenant Entra — **mesmo procedimento do `Calendars.ReadWrite`** da Fase 2: adicionar no registo da app + admin consent + atualizar `GRAPH_SCOPES` no `.env` de produção (a fonte de verdade) e no `config.py`/`.env.example` (lição da Fase 2). Sem este consent, as chamadas reais falham com reauth/consent; os **testes mockados** (FakeGraphClient) não são bloqueados.
2. Conta O365 **ligada** ao MCP (login concluído; `whoami` devolve a conta) e re-login após o consent para o token Graph passar a incluir os scopes de Teams.
3. Tenant/conta de teste com pelo menos um chat 1:1, um chat de grupo (com tópico) e mensagens (incl. uma de sistema), para o runbook de validação manual.

---

## 9. Faseamento e Definition-of-Done

1. **Esta entrega:** `teams_list_chats`, `teams_read_messages`, `teams_send_message_prepare`/`_confirm`; métodos Graph + mapeadores em `client.py` (`list_chats`, `list_chat_messages`, `send_chat_message`, e — se D1/D3 = criar — `get_or_create_one_on_one_chat`); registo em `server.py` + reforço de `instructions` (resolver nome antes do prepare; conteúdo não-confiável); reconciliar o playbook (§2.3 já usa estes nomes); FakeGraphClient estendido + E2E read/write + unit dos mapeadores; `docs/fase-3/estado-user-stories.md` + runbook (pelo QA).
2. **DoD:** 5 US com os critérios de §2; invariantes provados por contagem (prepare-não-escreve, idempotência, reauth graciosa, isolamento/TTL); sanitização + `content_is_untrusted` em todas as leituras; auditoria só-metadados em todos os envios; `ruff` limpo e `pytest` a passar. Decisões D1–D11 fechadas e refletidas no plano de implementação.
