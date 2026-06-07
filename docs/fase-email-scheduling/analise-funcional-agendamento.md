# Análise Funcional — Agendamento de envio de email
## Fase Email (extensão do Módulo Email, US-1.9 a US-1.11)

**Cliente:** Mobiweb
**Produto:** Servidor MCP de Office 365 — agendamento de envio de email (Microsoft Graph)
**Data:** 2026-06-07
**Estado:** Análise funcional para implementação. Reutiliza integralmente a arquitetura de identidade dual-plane, o modelo de aprovação em duas fases e a auditoria só-metadados já em produção na Fase 1. Decisões de produto fechadas pelo cliente (ver §2). Gates de validação no tenant real por fechar (ver §11).
**Referências:** [Análise Funcional v1.1](../analise-funcional-v1.1.md) (§2 identidade, §3 prepare/confirm, §4 segurança, §5 scopes, §7 RGPD) · [Fase 1 — estado das US](../fase-1/estado-user-stories.md) (US-1.1 a US-1.8; tratamento de anexos grandes em US-1.6) · [Fase 2 — estado das US](../fase-2/estado-user-stories.md) (fuso do mailbox, `MailboxSettings.Read`, header `Prefer: outlook.timezone`, degradação graciosa para UTC) · [Fase 3 — Teams](../fase-3/analise-funcional-teams.md) (profundidade/estilo de análise) · [Playbook do assistente](../../src/prompts/assistant-playbook.md)

> **Relação com a v1.1:** esta entrega **não introduz um módulo novo** — estende o Módulo Email da Fase 1 com três user stories (US-1.9 a US-1.11). Reutiliza **integralmente** a arquitetura dual-plane (§2), o two-phase approval server-side (§3), a auditoria só-metadados (§1.2/§7), a fronteira de conteúdo não-confiável (§4) e a reautenticação graciosa. **Não reinventa padrões** — segue `src/mcp_o365/tools/email.py` e `src/mcp_o365/graph/client.py`, reaproveitando o caminho `create_draft`→`send_draft` já existente para os anexos grandes (US-1.6).

---

## 1. Objetivo e enquadramento

A Fase 1 já permite **enviar, responder, reencaminhar, arquivar e eliminar** email com aprovação humana imposta server-side. Falta, no entanto, **agendar o envio de um email para uma data/hora futura** ("envia isto amanhã às 9h", "agenda para segunda de manhã"), bem como **listar** e **cancelar** os envios agendados ainda pendentes.

**Esta era uma falta de funcionalidade do produto, não uma limitação da API.** O Microsoft Graph suporta nativamente o envio diferido através de uma *extended property* MAPI do Exchange (`PidTagDeferredSendTime`) — a entrega é retida e processada **pelo próprio Exchange Online**, não pelo servidor MCP. Por isso o agendamento **não exige o servidor MCP vivo no momento do envio**: depois de o rascunho ser submetido com a propriedade, o Exchange entrega na hora marcada mesmo que a VPS esteja desligada. É a abordagem correta (nativa, fiável, sem agendador interno nem estado persistente do nosso lado) e foi a escolhida pelo coordenador.

**Dentro de âmbito:**
- Agendar o envio de um email (novo, com destinatários, assunto, corpo e anexos — incluindo anexos grandes >3MB) para um instante futuro (US-1.9).
- Listar os envios agendados ainda **pendentes** (US-1.10).
- Cancelar um envio agendado pendente antes de a hora chegar (US-1.11).
- Tradução de pedidos temporais para um instante absoluto no **fuso do mailbox** do utilizador (decisão 1 do cliente), com conversão interna para UTC ao gravar.

**Fora de âmbito (diferido, justificado):**
- **Reagendar** (alterar a hora de um agendamento existente) — na v1, reagendar = cancelar (US-1.11) + agendar de novo (US-1.9). Um `schedule_update` dedicado fica diferido (baixo valor face ao esforço; o caminho cancelar+reagendar é claro e sem ambiguidade).
- **Agendar respostas/reencaminhos** (`reply`/`forward` diferidos). A primeira entrega cobre o envio de mensagem nova (o caso pedido). O `reply`/`forward` diferido tem nuances próprias (o endpoint nativo `reply`/`forward` do Graph não aceita a extended property — exigiria `createReply`/`createForward` para obter um rascunho e só então aplicar a propriedade). Diferido.
- **Janelas de envio/"horário de expediente" automático** (ex.: "nunca enviar fora das 9h-18h" sem hora explícita) — política de produto, não pedida. Fora de âmbito.

---

## 2. Decisões de produto fechadas pelo cliente (não reabrir)

Registadas aqui para o Dev/QA as tratarem como "lei" (à imagem das decisões da Fase 2/3):

1. **Fuso horário:** as horas usam **sempre** o fuso do mailbox do utilizador (lido via `get_mailbox_timezone` / `MailboxSettings.Read`), exceto se o utilizador indicar **explicitamente** outro fuso. Nunca assumir UTC **na apresentação/interpretação**. Internamente, o instante de envio diferido é convertido para **UTC** ao gravar na propriedade Graph (a propriedade `SystemTime` é sempre UTC ISO 8601). Ver §6.
2. **Cancelamento:** tem de existir forma de **cancelar** um envio agendado pendente; **listar** os agendados pendentes também é desejável. Ver US-1.10/US-1.11 e §7.
3. **Mecanismo Graph:** envio diferido **nativo do Exchange** via a *single value extended property* `PidTagDeferredSendTime` (`SystemTime 0x3FEF`, valor UTC ISO 8601). Fluxo: criar rascunho (`POST /me/messages`) com a propriedade + chamar `send` no rascunho (`POST /me/messages/{id}/send`). **Não** via `POST /me/sendMail` direto (ver §3, nuance). Reaproveita o caminho draft→send da US-1.6.
4. **Aprovação:** agendar (US-1.9) e cancelar (US-1.11) seguem o **two-phase approval** server-side (`*_prepare` devolve resumo + `confirmation_token`; `*_confirm` executa). Listar (US-1.10) é leitura — sem aprovação.

---

## 3. Mecanismo Graph (com fontes oficiais)

O Exchange Online suporta envio diferido através da propriedade canónica MAPI **`PidTagDeferredSendTime`**, exposta no Microsoft Graph como uma **single value extended property** da mensagem. O cliente "marca" um rascunho com o instante de envio pretendido e submete-o; o Exchange retém a mensagem e entrega-a quando a hora for atingida.

### 3.1 A propriedade

| Campo | Valor |
|-------|-------|
| Propriedade canónica MAPI | `PidTagDeferredSendTime` |
| Property tag | `0x3FEF` |
| Tipo MAPI | `PT_SYSTIME` → no Graph o `graph_type` é **`SystemTime`** |
| `id` (formato `"{graph_type} {proptag}"`) | **`"SystemTime 0x3FEF"`** |
| `value` | Instante de envio em **UTC ISO 8601** (ex.: `"2026-06-10T09:00:00.000Z"` ou `"2026-06-10T09:00:00Z"`) |

**Semântica documentada:**
- A mensagem fica **retida na pasta Drafts** até o instante de envio diferido ser atingido — e só então é entregue pelo Exchange.
- Se o `value` for **anterior ao instante atual**, a mensagem é enviada **imediatamente** (relevante para a validação "no futuro" — ver §6.2: não confiar nesta tolerância, validar nós).
- O processamento é **server-side do Exchange** (transport), independente do servidor MCP.

### 3.2 O fluxo correto: draft → send (não `sendMail`)

A propriedade tem de ser definida na **criação do rascunho** e o rascunho é depois **submetido** — exatamente o caminho que a US-1.6 já usa para anexos grandes:

```
1) POST /me/messages
   { ...message..., "singleValueExtendedProperties": [ { "id": "SystemTime 0x3FEF", "value": "<UTC ISO8601>" } ] }
   -> devolve o rascunho com {id}

   (se houver anexos grandes >3MB: criar upload session por anexo + PUT dos bytes em chunks,
    exatamente como na US-1.6, ANTES do passo 2)

2) POST /me/messages/{id}/send
   -> 202 Accepted. O Exchange retém o rascunho e entrega na hora marcada.
```

> **Nuance importante (e porque não usar `sendMail`):** embora alguma documentação/exemplos da comunidade sugiram aplicar `singleValueExtendedProperties` no corpo de `POST /me/sendMail`, há relatos consistentes de que o `sendMail` **ignora/descarta** extended properties (ver Issue 2210 do SDK .NET nas fontes). O caminho **fiável e documentado** é **draft + send**. O coordenador fechou o caminho draft→send; esta análise confirma-o e recomenda **não** tentar o atalho `sendMail` com a propriedade (seria um envio imediato silencioso — falha de segurança, pois o utilizador esperava um envio diferido).

### 3.3 Métodos Graph a reutilizar/acrescentar em `client.py`

Já existem: `create_draft` (`POST /me/messages`), `send_draft` (`POST /me/messages/{id}/send`), `create_attachment_upload_session`, `upload_attachment_bytes`, `get_mailbox_timezone`, `list_messages`, `permanent_delete`. **Reaproveitar tudo isto.**

A acrescentar (mínimo):
- **Permitir `singleValueExtendedProperties` no `create_draft`** — o método já faz `POST /me/messages` com o objeto `message`; basta que o `message` montado pela tool possa incluir a coleção `singleValueExtendedProperties`. **Não é preciso novo endpoint** (é o mesmo `POST /me/messages`).
- **Listar rascunhos diferidos** (US-1.10) — `GET /me/mailFolders/drafts/messages` com `$filter`/`$expand` sobre a extended property (ver §7). Pode ser um método novo `list_deferred_drafts` ou um parâmetro de `list_messages` que aceite `expand`/`filter` da extended property na pasta `drafts`.
- **Cancelar** (US-1.11) — eliminar o rascunho diferido por id. Recomenda-se reutilizar `move_message(..., destination_id="deleteditems")` (soft, recuperável) ou `permanent_delete` consoante a decisão de cancelamento (ver §7.3). **Sem novo endpoint** se reutilizarmos os existentes.

### 3.4 Fontes

- [PidTagDeferredSendTime Canonical Property — Microsoft Learn](https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/pidtagdeferredsendtime-canonical-property) — definição da propriedade, tag `0x3FEF`, tipo `PT_SYSTIME`, semântica de retenção e "se anterior ao instante atual, envia já".
- [Create single-value extended property — Microsoft Graph v1.0 — Microsoft Learn](https://learn.microsoft.com/en-us/graph/api/singlevaluelegacyextendedproperty-post-singlevalueextendedproperties?view=graph-rest-1.0) — formato do `id` (`"{graph_type} {proptag}"`, `SystemTime` para `PT_SYSTIME`), e como incluir a coleção `singleValueExtendedProperties` ao criar o rascunho.
- [Create message — Microsoft Graph v1.0 — Microsoft Learn](https://learn.microsoft.com/en-us/graph/api/user-post-messages?view=graph-rest-1.0) e [message: send — Microsoft Graph v1.0 — Microsoft Learn](https://learn.microsoft.com/en-us/graph/api/message-send?view=graph-rest-1.0) — `POST /me/messages` (criar rascunho) e `POST /me/messages/{id}/send` (submeter; 202 Accepted).
- [Get singleValueLegacyExtendedProperty — Microsoft Graph v1.0 — Microsoft Learn](https://learn.microsoft.com/en-us/graph/api/singlevaluelegacyextendedproperty-get?view=graph-rest-1.0) — `$expand=singleValueExtendedProperties($filter=ep/id eq '...')` para ler/filtrar a propriedade (base da listagem, §7).
- [Delay message delivery with Graph API — Martin Machacek](https://martin-machacek.com/blogPost/b5ff46f7-b864-4075-bbfd-c7ef82370ed1) e [Send a delayed message — PnP Script Samples](https://pnp.github.io/script-samples/graph-delay-message-delivery/README.html) — exemplos end-to-end (draft com `"id": "SystemTime 0x3FEF"`, valor ISO 8601, retido em Drafts).
- [SendMail ignores extended MAPI properties specified as SingleValueExtendedProperties — Issue #2210, microsoftgraph/msgraph-sdk-dotnet](https://github.com/microsoftgraph/msgraph-sdk-dotnet/issues/2210) — confirma que `sendMail` **não** aplica a extended property; impõe o caminho draft→send.

> **Nota sobre as fontes:** a documentação oficial cobre exaustivamente **definir** a propriedade e **submeter** o rascunho. Não documenta de forma explícita o **cancelamento/recall** de uma mensagem diferida já submetida nem a sua localização exata após o `send` (Drafts vs Outbox vs fila de transporte). Isto é tratado honestamente como incerteza em §7.2 e marcado como **gate de validação no tenant real** (§11), no mesmo espírito do US-1.6 da Fase 1 (testes mockados passam; validação manual no tenant fica rastreada).

---

## 4. User Stories (US-1.9 a US-1.11)

Coerentes com o Módulo Email (continuação de US-1.1…US-1.8). Nomes de tool propostos, coerentes com os `email_*` existentes em `server.py`:
- US-1.9 → `email_schedule_prepare` / `email_schedule_confirm` (escrita, two-phase).
- US-1.10 → `email_list_scheduled` (leitura, sem aprovação).
- US-1.11 → `email_schedule_cancel_prepare` / `email_schedule_cancel_confirm` (escrita, two-phase).

| US | Título | Tipo |
|----|--------|------|
| **US-1.9** | Agendar envio de email para uma data/hora | Escrita (prepare/confirm) |
| **US-1.10** | Listar envios agendados pendentes | Leitura |
| **US-1.11** | Cancelar um envio agendado pendente | Escrita (prepare/confirm) |

### US-1.9 — Agendar envio de email

Como utilizador, quero pedir "envia este email amanhã às 9h" para que o Exchange entregue a mensagem no instante indicado, sem o servidor MCP precisar de estar ligado nessa altura.

**Critérios de aceitação:**

1. **[AC-WRITE]** `email_schedule_prepare` valida e devolve **resumo + `confirmation_token`** e **NÃO escreve no Graph** (não cria rascunho, não submete). `email_schedule_confirm` só executa com token válido, não expirado e não usado (idempotency key). Provado por contagem (`create_draft`/`send_draft` a **0** após o `prepare`).
2. **Parâmetros** idênticos aos do envio (US-1.3): `to` (obrigatório), `subject_line`, `body`, `body_type`, `cc`, `bcc`, `attachments`, `message_meta` — **mais** o instante de envio: `send_at` (ISO 8601, no fuso do mailbox salvo se trouxer offset/fuso explícito) e, opcionalmente, `timezone` (fuso explícito que o utilizador indicou). Sem `to` → `error`. Sem `send_at` → `error`.
3. **Resolução do fuso e cálculo do instante UTC** (decisão 1): o `prepare` resolve o fuso do mailbox **uma vez** (best-effort, via `get_mailbox_timezone`; degradação graciosa para UTC se `MailboxSettings.Read` faltar — igual à Fase 2, nunca derruba a sessão). O `send_at` é interpretado nesse fuso (ou no `timezone` explícito) e **convertido para UTC ISO 8601** para a propriedade. O resumo declara a hora **no fuso do utilizador** e o fuso usado (ex.: "agendar para 10/06/2026 09:00 (Hora de Lisboa)").
4. **Validação temporal** (§6.2): `send_at` tem de estar **no futuro** com **margem mínima** (ex.: ≥ 2 minutos) e abaixo de um **limite superior** razoável (ex.: ≤ 1 ano). Passado, margem insuficiente, limite excedido ou `send_at` não-parseável → `error` orientador (**sem token**), antes de qualquer escrita.
5. **Resumo (`prepare`)** declara, sem PII de conteúdo: nº de destinatários e **domínios** (via `_domains`), assunto, presença de anexos (e se inclui anexos grandes → via upload session), e o **instante de envio no fuso do utilizador**. Texto-tipo: *"Agendar email para 3 destinatário(s) (domínios: mobiweb.pt), assunto 'X', envio em 10/06/2026 09:00 (Hora de Lisboa)."*
6. **[AC-WRITE]** O `confirm` executa o caminho **draft→send**: `create_draft` com a `message` **+** `singleValueExtendedProperties` (`"SystemTime 0x3FEF"` = UTC) e, se houver anexos grandes, upload session por anexo (reutiliza US-1.6) **antes** do `send_draft`. Devolve o `id`/referência do rascunho diferido (para permitir cancelar depois) e a hora de envio. Idempotência: replay do token → `idempotent_replay=true` **sem** segundo draft/send.
7. **Reautenticação graciosa** em ambas as fases: qualquer `invalid_grant`/401/403 → `reauth_required`; no `confirm`, em falha de refresh o token **não é consumido** (repetível após re-login). A leitura acessória do fuso **nunca** derruba a sessão.
8. **Auditoria só-metadados** no `confirm`: `action="email.schedule"`, `recipients_count`, `extra` = `{large_attachments, send_at_utc, deferred=true}` — **nunca** o corpo, nem endereços em claro (só contagem/domínios via o padrão já existente). Ver §8.
9. **Anexos grandes (>3MB)** coexistem com o envio diferido (§5/§6.4): o rascunho leva os anexos inline (≤3MB) e a extended property na criação; os grandes seguem por upload session; só depois o `send_draft`. O resumo marca `large_attachments=true`.

### US-1.10 — Listar envios agendados pendentes

Como utilizador, quero ver os emails que tenho agendados e ainda não enviados, com hora e destinatário, para poder controlar/cancelar.

**Critérios de aceitação:**

1. **Leitura, sem aprovação.** `email_list_scheduled` não exige token; não escreve.
2. **Fonte de verdade = o mailbox** (decisão de arquitetura, §7.1): lista os **rascunhos** que têm a extended property `PidTagDeferredSendTime` definida e cujo instante é **futuro** (ainda pendentes). `GET /me/mailFolders/drafts/messages` com `$filter=singleValueExtendedProperties/any(ep: ep/id eq 'SystemTime 0x3FEF')` e `$expand=singleValueExtendedProperties($filter=ep/id eq 'SystemTime 0x3FEF')` para trazer o valor da hora.
3. **Por item devolve:** `id` (do rascunho — usado para cancelar), `subject`, destinatários como **contagem + domínios** (não endereços em claro no log; na resposta ao utilizador os destinatários podem ser mostrados — é o seu próprio rascunho), `send_at` apresentado **no fuso do mailbox** (convertendo do UTC guardado) e `send_at_utc`. Conteúdo (corpo) **não** é devolvido por defeito (minimização; é uma listagem).
4. **Paginação consciente** seguindo o padrão do `email_search` (auto-paginar até um teto `_MAX_FETCH_ALL`, ou devolver 1ª página + `has_more` — recomenda-se devolver todos, pois o nº de agendamentos pendentes é tipicamente pequeno).
5. **Fronteira de não-confiança:** o assunto/preview, sendo conteúdo de mensagem, é tratado como **não-confiável** (`content_is_untrusted=true`) e sanitizado se HTML — mesmo sendo um rascunho do próprio (defesa em profundidade: pode ter sido criado por outra via).
6. **Reautenticação graciosa** coberta (`reauth_required`); a leitura do fuso é best-effort (degradação para UTC).

> **Nota de robustez (gate, §11):** o filtro por extended property em mensagens pode não suportar comparação de `ep/value` (data) de forma fiável em todos os tenants. **Mitigação:** filtrar pela **presença** da propriedade (`any(ep: ep/id eq 'SystemTime 0x3FEF')`) e, se necessário, **filtrar o "ainda futuro" client-side** comparando o `send_at_utc` expandido com o instante atual. Validar no tenant real.

### US-1.11 — Cancelar um envio agendado pendente

Como utilizador, quero cancelar um agendamento antes de a hora chegar, para que o email não seja enviado.

**Critérios de aceitação:**

1. **[AC-WRITE]** `email_schedule_cancel_prepare` valida e devolve **resumo + `confirmation_token`** e **NÃO escreve**. `email_schedule_cancel_confirm` só executa com token válido. Idempotência por token (replay → `idempotent_replay=true` sem segunda eliminação).
2. **Parâmetro:** `message_id` (o `id` do rascunho diferido, obtido de US-1.10 ou do retorno de US-1.9).
3. **`prepare`** confirma (best-effort) que o `message_id` corresponde a um rascunho **com a extended property** (ainda pendente) e monta um resumo: *"Cancelar o envio agendado para 10/06/2026 09:00 (Hora de Lisboa), assunto 'X', N destinatário(s)."* Se o rascunho já não tiver a propriedade (já foi enviado / não é um diferido) → `error` orientador **sem token** (ver §7.2, gate de cancelabilidade).
4. **[AC-WRITE]** O `confirm` **elimina o rascunho** por id — recomenda-se **soft delete** (mover para Itens Eliminados, recuperável) por defeito, à imagem do US-1.8; eliminação permanente do rascunho fica como opção reforçada (não default). Reutiliza `move_message`/`permanent_delete`. Devolve confirmação.
5. **Reautenticação graciosa** em ambas as fases; token não consumido em falha de refresh no `confirm`.
6. **Auditoria só-metadados:** `action="email.schedule_cancel"`, `target=message_id`, `extra` = `{permanent: false}` — sem corpo nem endereços em claro.
7. **Cancelabilidade é um gate de validação no tenant real** (§7.2/§11): a eliminação do rascunho **antes** da hora de envio deve impedir a entrega, mas isto **tem de ser confirmado no tenant** (janela de corrida perto da hora de envio; comportamento do transport quando o item sai de Drafts). Honestamente declarado como incerteza até validação manual.

---

## 5. Anexos grandes (>3MB) e envio diferido

O envio diferido e o upload de anexos grandes **coexistem** sem conflito, porque ambos assentam no mesmo caminho **draft→send** (já implementado e testado na US-1.6):

```
create_draft({ ...message (anexos inline <=3MB), singleValueExtendedProperties:[deferred] })
  -> por cada anexo grande: createUploadSession + upload_attachment_bytes (chunks 320 KiB, sem Bearer)
  -> send_draft   (o Exchange retém o rascunho — com anexos completos — até send_at)
```

A extended property é definida na **criação** do rascunho (passo 1); os anexos grandes são carregados no rascunho **antes** do `send_draft`. Não há ordem alternativa nem endpoint extra. **Gate (§11):** validar no tenant real um agendamento com anexo >3MB (a US-1.6 já tem este envio normal marcado como ⬜ na validação manual — o caso diferido herda o mesmo gate).

---

## 6. Tratamento de fuso horário e validações

### 6.1 Fuso (decisão 1)

- O `prepare` resolve o fuso do mailbox **uma vez** via `get_mailbox_timezone` (best-effort; `MailboxSettings.Read`). Se o utilizador indicou um **fuso explícito** (`timezone`), esse prevalece.
- O `send_at` do utilizador é interpretado nesse fuso e **convertido para UTC** para a propriedade `SystemTime` (sempre UTC ISO 8601, ex.: `2026-06-10T09:00:00.000Z`).
- **Apresentação:** o resumo do `prepare`, a listagem (US-1.10) e o resumo do cancelamento (US-1.11) mostram a hora **no fuso do utilizador** — nunca UTC cru. Guarda-se/auditam-se ambos quando útil (`send_at_utc`).
- **Degradação graciosa:** sem `MailboxSettings.Read`, assume-se UTC para o cálculo **e o resumo declara-o explicitamente** ("fuso do mailbox indisponível; a interpretar como UTC") — para o utilizador poder corrigir. A leitura do fuso nunca derruba a sessão (lição da Fase 2, `_resolve_tz` best-effort).

### 6.2 Validação temporal

| Regra | Comportamento |
|-------|---------------|
| `send_at` não-parseável | `error` orientador, sem token |
| `send_at` no passado (ou < margem mínima do futuro) | `error` — **não** confiar na tolerância do Exchange ("se anterior ao agora, envia já"): seria um envio imediato inesperado. Validar nós. |
| Margem mínima | ex.: **≥ 2 minutos** no futuro (evita corrida com o relógio/latência); valor a confirmar no contrato. |
| Limite superior | ex.: **≤ 1 ano** (rascunhos diferidos a anos são provavelmente erro; valor a confirmar). |

### 6.3 UX do assistente (tradução de pedidos)

O assistente traduz expressões naturais ("envia amanhã às 9h", "agenda para segunda de manhã", "daqui a duas horas") para um `send_at` ISO 8601 **no fuso do mailbox** (ou no fuso que o utilizador nomear), e **confirma a hora absoluta calculada** com o utilizador antes de chamar `email_schedule_prepare` — e novamente no `confirm` (resumo humano com a hora no fuso local). Ambiguidades ("segunda" = qual?) → o assistente esclarece antes de preparar. As tools recebem o `send_at` **já resolvido** (o servidor não faz parsing de linguagem natural — só de ISO 8601), espelhando o padrão "tools recebem valores resolvidos" do Calendário/Teams.

### 6.4 Interação com anexos grandes

Coberto em §5 — coexistem; sem ordem alternativa.

---

## 7. Listar/cancelar — fonte de verdade (análise e recomendação)

### 7.1 Mailbox vs registo local — **recomendação: o mailbox é a única fonte de verdade**

| Critério | Mailbox (rascunhos diferidos) | Registo local na VPS |
|----------|-------------------------------|----------------------|
| Drift de estado | **Nenhum** — quem retém/envia é o Exchange; o que lá está é a verdade | **Risco alto** — o utilizador pode apagar o rascunho pelo Outlook; o nosso registo fica obsoleto |
| Servidor MCP desligado no envio | Irrelevante (Exchange entrega) | Irrelevante para o envio, mas o registo desatualiza |
| Multi-conta / multi-dispositivo | Consistente (o mailbox é partilhado) | Teria de sincronizar por conta |
| RGPD | Sem dados extra na VPS | Mais um sítio com metadados de email |
| Complexidade | Baixa (reutiliza `list_messages` + filtro) | Alta (persistência, sincronização, limpeza) |

**Recomendação: usar o mailbox como fonte de verdade** — listar os rascunhos com a extended property e cancelar eliminando o rascunho por id. **Sem registo local, sem agendador, sem estado a sincronizar.** É a solução mais simples e a única sem drift. (Coerente com o princípio do projeto: o Exchange faz o trabalho; nós não duplicamos estado.)

### 7.2 Cancelabilidade — incerteza honesta (GATE)

A documentação oficial **não** descreve explicitamente como cancelar/recall de uma mensagem **já submetida** com `PidTagDeferredSendTime`. O modelo documentado é: o rascunho fica em **Drafts** até à hora. A hipótese de trabalho (e a abordagem recomendada) é: **eliminar o rascunho por id antes da hora de envio cancela o envio** (deixa de haver item para o transport processar). Há, contudo, incertezas a validar no tenant:
- **Onde fica exatamente o item após o `send_draft`** — a fonte indica "Drafts até à hora", mas convém confirmar se, perto da hora, transita para Outbox/fila de transporte (janela em que eliminar pode já não impedir).
- **Janela de corrida** perto do `send_at`: cancelar segundos antes pode falhar se o transport já pegou o item.

**Tratamento (padrão do projeto):** os testes automáticos cobrem o fluxo (mockado: cria draft com propriedade, lista, elimina) e **passam**; a **cancelabilidade real** fica **⬜ pendente de validação manual no tenant** e rastreada como gate (§11), à imagem do US-1.6 da Fase 1. O `prepare` do cancelamento devolve `error` orientador se o rascunho já não estiver pendente (best-effort), e a documentação da tool declara a limitação ("cancelar muito perto da hora pode já não impedir o envio").

### 7.3 Soft vs hard delete no cancelamento

Recomenda-se **soft delete** (mover para Itens Eliminados, recuperável) por defeito — coerente com US-1.8; o utilizador recupera o rascunho se cancelou por engano. Hard delete (`permanent_delete`) fica como opção reforçada, não default.

---

## 8. Segurança, RGPD e auditoria

- **Fronteira de prompt injection inalterada (§4 v1.1):** o agendamento **não muda** a fronteira. O `body`/`subject` continuam **conteúdo não-confiável** — o assistente nunca age por instruções vindas de dentro do conteúdo; só por intenção direta do utilizador, com resumo real no `confirm`. A listagem (US-1.10) sanitiza assunto/preview e marca `content_is_untrusted=true`.
- **Auditoria só-metadados** (`log_audit`, `subject_hash`), retenção 12 meses com purga (v1.1 §7):
  - `email.schedule` (US-1.9 confirm): `recipients_count`, `extra={large_attachments, send_at_utc, deferred:true}`.
  - `email.schedule_cancel` (US-1.11 confirm): `target=message_id`, `extra={permanent:false}`.
  - **Nunca** o corpo, nem o assunto em claro, nem endereços (no máximo contagem/domínios, via o padrão `_domains`). Sem `subject_hash` em `extra` (regra A1 do projeto — só o de topo, identidade).
- **Soberania de dados:** a mesma tensão da v1.1 §7 — o corpo do email é enviado ao modelo Claude na composição; o agendamento não acrescenta tratamento novo além de metadados (hora de envio) na VPS/logs.
- **Isolamento por `subject`:** cada utilizador agenda/lista/cancela apenas sobre o seu próprio mailbox (token delegado); sem acesso cruzado.

---

## 9. Scopes Microsoft Graph

| Operação | Endpoint | Scope |
|----------|----------|-------|
| Criar rascunho diferido | `POST /me/messages` (com `singleValueExtendedProperties`) | `Mail.ReadWrite` |
| Upload session de anexo grande | `POST /me/messages/{id}/attachments/createUploadSession` + PUT | `Mail.ReadWrite` |
| Submeter rascunho | `POST /me/messages/{id}/send` | `Mail.Send` |
| Listar rascunhos diferidos | `GET /me/mailFolders/drafts/messages?$filter=…&$expand=…` | `Mail.Read` (ou `Mail.ReadWrite`) |
| Cancelar (mover/eliminar rascunho) | `POST /me/messages/{id}/move` ou `permanentDelete` | `Mail.ReadWrite` |
| Fuso do mailbox | `GET /me/mailboxSettings` | `MailboxSettings.Read` (best-effort; já concedido na Fase 2) |

**Conclusão:** os scopes **já concedidos** chegam — `Mail.Read` + `Mail.Send` + `Mail.ReadWrite` (Fase 1) cobrem agendar/listar/cancelar; `MailboxSettings.Read` (Fase 2) cobre o fuso. **Não é preciso novo admin consent.** (Confirmar na validação real que o token efetivo inclui estes scopes, mas não há scope novo a pedir.)

---

## 10. Invariantes a herdar (Fase 1 → esta entrega)

Sem exceções; provados por contagem de chamadas (FakeGraphClient):

1. **`prepare` NÃO escreve:** `create_draft`/`send_draft`/`move_message` a **0** após o `prepare` (agendar e cancelar).
2. **Token fresco no `confirm` + idempotência:** o token é idempotency key; replay → `idempotent_replay=true` sem segundo draft/send/delete (evita **duplo agendamento** e **duplo envio** — risco real).
3. **TTL / isolamento por `subject`:** token expirado → `expired`; token de outro subject → `error`.
4. **Reautenticação graciosa:** `invalid_grant`/401/403 → `reauth_required`; no `confirm`, token não consumido em falha de refresh.
5. **Conteúdo não-confiável:** assunto/preview da listagem sanitizados + `content_is_untrusted=true`.
6. **Auditoria só-metadados** em todas as escritas.
7. **Fuso lido 1×/pedido, best-effort:** nunca derruba a sessão (lição da Fase 2).

---

## 11. Riscos e gates de validação no tenant real

| Risco / incerteza | Impacto | Mitigação / gate |
|---|---|---|
| **Cancelabilidade real** de um rascunho diferido já submetido (§7.2) | **Alto** (a funcionalidade de cancelar depende disto) | **GATE:** validar no tenant que eliminar o rascunho antes do `send_at` impede a entrega; mapear a janela de corrida perto da hora. Testes auto passam; validação manual ⬜. |
| **Localização do item** após `send_draft` (Drafts vs Outbox/transport) (§3.1/§7.2) | Médio-Alto | **GATE:** confirmar onde reside e quando deixa de ser cancelável. Documentar a janela na descrição da tool. |
| **Filtro por extended property** (`$filter`/`$expand`) na pasta drafts não fiável em todos os tenants (US-1.10) | Médio | Filtrar por **presença** da propriedade + filtrar "futuro" **client-side** sobre o valor expandido. **GATE:** validar a query no tenant. |
| **`sendMail` ignora a extended property** (§3.2) | Alto (se alguém atalhar) | Usar **sempre** draft→send; não usar `sendMail` com a propriedade. Coberto por design + teste. |
| **Envio imediato indesejado** se `send_at` no passado/sem margem (§6.2) | Médio | Validar futuro + margem mínima no `prepare`, **antes** de qualquer escrita; recusar sem token. |
| **Anexo grande >3MB diferido** não exercido no real (§5) | Médio | **GATE:** herda o gate do US-1.6 (envio >3MB real ⬜). Validar um agendamento com anexo grande. |
| **Duplo agendamento/envio** (retry/re-chamada do LLM) | Médio-Alto | Idempotency key = `confirmation_token`; replay não re-cria nem re-envia. |
| **Prompt injection** via corpo do email a agendar | Alto | Fronteira inalterada (v1.1 §4); só intenção direta; risco residual aceite. |
| **Fuso do mailbox indisponível** (`MailboxSettings.Read`) | Baixo-Médio | Degradação graciosa para UTC **declarada no resumo**; best-effort, nunca derruba a sessão. |

> **Padrão de rastreio (igual ao US-1.6):** todos os gates acima entram em [estado-user-stories.md](estado-user-stories.md) como **Validação manual ⬜** mesmo quando os testes automáticos (mockados) estão ✅. A validação manual no tenant/VPS reais é responsabilidade do cliente (runbook), à imagem das Fases 1 e 2.

---

## 12. Definition of Done (resumo)

- 3 US (US-1.9, US-1.10, US-1.11) com os critérios de §4; AC-WRITE = par prepare/confirm para US-1.9 e US-1.11.
- `client.py`: `create_draft` aceita `singleValueExtendedProperties`; método de listagem de rascunhos diferidos; cancelamento reutiliza `move_message`/`permanent_delete`. Reutiliza o caminho da US-1.6 para anexos grandes.
- `email.py`: `run_email_schedule_prepare`/`_confirm`, `run_email_list_scheduled`, `run_email_schedule_cancel_prepare`/`_confirm` — seguindo o estilo das tools existentes.
- `server.py`: registo das tools `email_schedule_*`, `email_list_scheduled`, `email_schedule_cancel_*` com descrições que instruem o LLM (resolver a hora no fuso do mailbox a montante; conteúdo não-confiável; cancelar perto da hora pode falhar).
- Invariantes de §10 provados por contagem; auditoria só-metadados (`email.schedule`/`email.schedule_cancel`); sanitização + `content_is_untrusted` na listagem; `ruff` limpo e `pytest` a passar.
- [estado-user-stories.md](estado-user-stories.md) atualizado; gates de §11 rastreados como validação manual ⬜.
