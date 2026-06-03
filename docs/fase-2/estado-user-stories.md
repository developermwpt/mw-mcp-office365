# Fase 2 — Módulo Calendário: estado das user stories

> **Pré-requisito da validação real (R1).** Tal como o `Mail.*` na Fase 1, a escrita de
> calendário precisa de **admin consent** do scope `Calendars.ReadWrite` no tenant (cobre
> leitura, escrita e `getSchedule` delegado). **Não bloqueia os testes mockados** (todos a
> passar), mas bloqueia a validação manual no tenant/VPS reais — ver `runbook-validacao-manual.md`.
>
> **Scope adicional (R1b).** Ler o **fuso do mailbox** (D1) exige `MailboxSettings.Read` — um
> scope distinto do `Calendars.*`. Sem ele o Graph devolve 403 nessa leitura; o servidor
> **degrada graciosamente para UTC** (a leitura do fuso é best-effort e nunca derruba a sessão
> — ver melhoria pós-deploy abaixo), mas só com `MailboxSettings.Read` concedido as horas saem
> no fuso real do utilizador. Já incluído em `GRAPH_SCOPES` (`.env`/`config.py`).

## Legenda

- ✅ feito · ⬜ pendente
- **Testado (automático):** coberto por testes unit/integração com Graph/Entra mockados.
- **Validação manual (tenant real):** execução no tenant/VPS reais — responsabilidade do
  cliente (ver runbook). Pendente em todas as US enquanto não houver admin consent de
  `Calendars.ReadWrite` e acesso ao tenant real.

## Tabela de estado

| US | Descrição curta | Implementado | Testado (auto) | Validação manual | Notas |
|----|-----------------|:---:|:---:|:---:|-------|
| US-2.1 | Listar eventos (intervalo, auto-paginação, fuso, body sanitizado) | ✅ | ✅ | ✅¹ | Auto-pagina TODO o intervalo seguindo `@odata.nextLink` (D5), sem perguntar; `auto_fetched_all=true`. Fuso do mailbox lido 1×/pedido (D1, best-effort) e devolvido em `timezone`. Cada evento expõe `responseStatus` (a resposta do próprio) e `isOrganizer` — permite "quais por aceitar?" na própria listagem. Ocorrências de séries vêm expandidas (`isRecurring`). `bodyPreview` sanitizado + `content_is_untrusted`. Teto `_MAX_FETCH_ALL` → `truncated_at`/`fetched_all=false`. |
| US-2.2 | Verificar disponibilidade (`getSchedule`) | ✅ | ✅ | ⬜ | `POST /me/calendar/getSchedule` (D2); o **próprio é sempre incluído** (via `/me`), dedup case-insensitive dos emails. `attendees` pode ser vazio (só o próprio). Horas no fuso do mailbox. |
| US-2.3 | Criar evento (prepare/confirm) | ✅ | ✅ | ⬜ | D6: sem `location` → link Teams; com `location` → presencial sem link — o resumo declara sempre. D3: resumo declara "Notifica N participante(s) (domínios: …)". prepare lê o fuso mas **não cria** (`create_event` a 0); confirm cria 1×; replay idempotente; auditoria `calendar.create` (só-metadados, `subject_hash`+`online`). |
| US-2.4 | Editar/reagendar (prepare/confirm; recorrência → clarification) | ✅ | ✅ | ⬜ | D4: recorrente sem `scope` → `needs_clarification` (sem token, `update_event` a 0). `scope='occurrence'` → PATCH ao próprio id; `scope='series'` → PATCH ao `seriesMasterId`. Só campos não-`None` entram em `changes`. Idempotência; auditoria `calendar.update` (`scope`). |
| US-2.5 | Cancelar evento (prepare/confirm) | ✅ | ✅ | ⬜ | Só o **organizador** cancela; não-organizador → `error` orientando para `decline` (R3, antes de qualquer escrita). Recorrente sem `scope` → clarification. **Mensagem de cancelamento (melhoria 2026-06-03):** se `message_choice_confirmed=false` → `needs_clarification` (mensagem própria / **sugestão a aceitar antes** / nenhuma); sem token, sem cancel. Resumo declara N participantes notificados + "Alto impacto". Idempotência; auditoria `calendar.cancel`. |
| US-2.6 | Responder a convite (accept/decline/tentative) | ✅ | ✅ | ⬜ | D7: prepare lê o estado atual (`responseStatus`) e **declara a transição** ("Já tinha Aceitado; vai mudar para Recusado…"); bloqueia se o subject for o **organizador** (sem token). `response` inválida → `error`. **Recusar (decline)**: se `message_choice_confirmed=false` → `needs_clarification` a perguntar se quer enviar mensagem ao organizador e qual (sem token); repetir com `comment` (com mensagem), `comment=''` (sem mensagem, notifica) ou `notify_organizer=false` (sem notificar). Idempotência; auditoria `calendar.respond` (`response`+`previous`+`notified`). |

> ¹ US-2.1 validada no tenant real em 2026-06-03 (listagem de 7 eventos com Teams e fuso de
> Lisboa) após as correções pós-deploy abaixo. As restantes US (2.2–2.6) continuam ⬜ até
> validação manual no tenant real.

## Detalhe por user story

### US-2.1 — `calendar_list_events` (leitura)

Resolve o fuso do mailbox (D1) **uma vez por pedido**, lê a 1ª página de `calendarView` e
**auto-pagina** todo o intervalo seguindo `@odata.nextLink` (D5, mesmo loop do `email_search`
mas SEM perguntar), com o teto de segurança `_MAX_FETCH_ALL=1000`. Cada evento é mapeado por
`_map_event_summary`; o `bodyPreview` é conteúdo NÃO-confiável e passa por `sanitize_html`. A
resposta inclui `timezone`, `auto_fetched_all`, `has_more`, `fetched_all` e
`content_is_untrusted`. `start`/`end` obrigatórios (`error` se faltar). Sem conta →
`reauth_required` e o Graph não é tocado.

### US-2.2 — `calendar_check_availability` (leitura)

`getSchedule` chamado **uma vez** com `schedules = [próprio, *attendees]`; o email do próprio
vem de `/me` (`userPrincipalName`). Dedup case-insensitive (não duplica o próprio se vier nos
`attendees`). Devolve, por pessoa, `availabilityView` + `scheduleItems`. `attendees` vazio →
só o próprio.

### US-2.3 — `calendar_create` (prepare/confirm)

Decisão Teams (D6): `online_meeting = online if online is not None else (location is None)`. O
payload Graph montado no prepare é guardado no pending op (não escrito). O resumo declara a
contagem de participantes notificados (com domínios, via `_domains`) e a frase Teams correta.
O confirm cria via `create_event`, devolve `event_id`/`web_link` e audita `calendar.create`.

### US-2.4 — `calendar_update` (prepare/confirm)

O prepare lê o evento (`get_event`) para detetar recorrência (`_is_recurring`). Recorrente sem
`scope` → `needs_clarification` (esta ocorrência vs série) sem token e sem PATCH. O alvo do
PATCH depende do scope: `occurrence` → próprio id; `series` → `seriesMasterId`. Só os campos
fornecidos entram em `changes`.

### US-2.5 — `calendar_cancel` (prepare/confirm)

O prepare lê o evento e o email do próprio: se o organizador ≠ próprio → `error` orientando
para `decline` (R3), antes de qualquer escrita. Recorrência tratada como no update. O resumo
declara o impacto e a notificação.

**Melhoria 2026-06-03 — mensagem de cancelamento.** Como cancelar **notifica sempre** os
participantes, se o utilizador ainda não decidiu (`message_choice_confirmed=false`) o prepare
devolve `needs_clarification` com 3 opções: mensagem **própria** (`comment='<texto>'`), pedir
uma **sugestão** (o assistente propõe um texto e o utilizador tem de **aceitar/ajustar antes**
— nunca se cancela com sugestão não aprovada), ou **sem mensagem** (`comment=''`). Só após a
escolha (com `message_choice_confirmed=true`) é emitido o token. Cobertura:
`test_cancel_sem_escolha_pede_mensagem` e `test_cancel_com_mensagem_confirmada`.

### US-2.6 — `calendar_respond` (prepare/confirm)

`response ∈ {accept, decline, tentative}` (senão `error`). O prepare lê o evento; se o
organizador for o próprio → `error` (sem token). Lê o `responseStatus` atual e declara a
transição PT no resumo. O confirm responde via `respond_event` e audita `previous`+`response`.

**Melhoria 2026-06-03 — mensagem na recusa.** Ao **recusar**, se o utilizador ainda não
decidiu (`message_choice_confirmed=false`), o prepare devolve `needs_clarification` a
perguntar se quer enviar mensagem ao organizador e qual o texto — **sem emitir token nem
responder**. As 3 opções devolvidas: recusar **com** mensagem (`comment='<texto>'`), **sem**
mensagem mas notificando (`comment=''`), ou **sem notificar** o organizador
(`notify_organizer=false` → `send_response=false`). `accept`/`tentative` não disparam a
pergunta e notificam sempre. O `send_response` escolhido é guardado no payload e auditado em
`notified`. Cobertura: `test_respond_decline_*` em `test_calendar_write_e2e.py`.

## Correções e melhorias pós-deploy (2026-06-03, validação no tenant real)

Após o primeiro deploy da Fase 2, a validação real expôs três pontos, todos corrigidos e
re-deployados no mesmo dia:

1. **Scope `Calendars.*` em falta no `.env` de produção** (commit que precedeu os abaixo). O
   `.env` definia `GRAPH_SCOPES` explicitamente (sobrepondo o default do `config.py`) sem os
   scopes de calendário. Adicionados `Calendars.Read`/`Calendars.ReadWrite`. **Lição:** o
   `.env` de produção é a fonte de verdade dos scopes; alterar o default do `config.py` não
   chega — alinhar sempre os dois (e o `.env.example`).

2. **Fuso (D1) deixou de derrubar a sessão — `_resolve_tz` best-effort** (commit `f491314`).
   `GET /me/mailboxSettings` exige `MailboxSettings.Read` (≠ `Calendars.*`); sem esse scope o
   Graph devolvia 403. Como `_resolve_tz` passava por `call_graph`, o 403 escalava a
   refresh→retry→reauth e **marcava a conta inteira como expirada** — derrubando também o
   email. Corrigido: `_resolve_tz` resolve o token e chama `get_mailbox_timezone` diretamente,
   apanhando `UpstreamAuthError`/`ReauthRequired` e devolvendo `None` (fallback UTC). Uma
   leitura acessória **nunca mais** marca a sessão como expirada. Adicionado
   `MailboxSettings.Read` aos scopes. Regressão:
   `test_fuso_403_nao_derruba_sessao_nem_falha_listagem`.

3. **`responseStatus`/`isOrganizer` na listagem** (commit `436f50b`). `_map_event_summary` não
   expunha a resposta do próprio ao evento, pelo que não era possível responder a "quais
   eventos estão por aceitar?" sem abrir cada um. O `calendarView` já devolve esses campos por
   defeito (não há `$select` a restringir) — faltava mapeá-los. Acrescentados ao resumo +
   descrição da tool a explicar o filtro dos pendentes (`responseStatus ∈ {notResponded,none}`
   e `isOrganizer=false`). Teste do mapper atualizado.

## Garantias transversais (verificadas por testes)

- **prepare NÃO escreve:** em todas as 4 escritas, `create_event`/`update_event`/
  `cancel_event`/`respond_event` ficam a **0** após o prepare (provado por contagem).
- **Idempotência:** replay de um token consumido devolve `idempotent_replay=true` **sem
  re-executar** (a operação real fica a 1).
- **TTL / isolamento:** token expirado → `expired`; token de outro `subject` → `error`
  (`ConfirmationNotFound`). Em ambos, a escrita real fica a 0.
- **Reautenticação graciosa:** falha de refresh no confirm → `reauth_required`, sem chamar o
  Graph para escrita; o token de confirmação **não é consumido** (repetível após re-login).
- **Recorrência → clarification:** editar/cancelar série recorrente sem `scope` devolve
  `needs_clarification` sem token e sem PATCH/cancel.
- **Fuso lido 1×:** `get_mailbox_timezone` chamado exatamente uma vez por pedido e devolvido
  em `timezone`; o header `Prefer: outlook.timezone` é injetado nas leituras.
- **Teto de paginação:** atingido `_MAX_FETCH_ALL`, a listagem trunca com `truncated_at` e
  `fetched_all=false` (não segue mais páginas).
- **Auditoria só-metadados:** cada escrita emite `event=audit` com `subject_hash` (nunca o
  assunto em claro), `target` (event_id / seriesMasterId), `recipients_count` e extras seguros
  (`online`, `scope`, `response`/`previous`). Nunca emails em claro nem o corpo.
- **Sanitização:** o `bodyPreview` dos eventos é sanitizado antes de chegar ao modelo;
  `content_is_untrusted=true` sempre presente nas leituras.

## Onde estão os testes

- Unit: `tests/unit/test_graph_calendar_client.py` — mapeadores (`_map_event_summary`/
  `_map_event_detail`/`_is_recurring`), header `Prefer`, body do `getSchedule`, roteamento de
  `respond_event` (accept/decline/tentativelyAccept).
- Integração (tools ponta-a-ponta): `tests/integration/test_calendar_read_e2e.py` (US-2.1,
  US-2.2) e `tests/integration/test_calendar_write_e2e.py` (US-2.3–2.6 + transversais), com o
  `FakeGraphClient` estendido em `tests/integration/fake_graph.py`.

Correr: `python -m pytest -q` · lint: `python -m ruff check src tests`.
