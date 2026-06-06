# Fase 3 — Módulo Teams (chats): estado das user stories

> **Plano APROVADO COM RESSALVAS pelo coordenador.** A revisão crítica independente
> ([revisao-coordenador.md](revisao-coordenador.md)) deu **APROVADO COM RESSALVAS**, sem
> bloqueadores, com dois achados Maiores (A1 e A2) a fechar antes das escritas. **Ambos já
> estão no código:** A1 — a auditoria `teams.send` NÃO leva `subject_hash` em `extra` (o
> `log_audit` já emite o `subject_hash` da *identidade* no campo de topo; um segundo
> sobrescreveria-o); A2 — o resumo de `teams_send_message_prepare` é montado a partir de
> `get_chat(chat_id)` (leitura pontual robusta a >50 chats), não de um match sobre `list_chats`.

> **Pré-requisito da validação real (R1) — admin consent CONCEDIDO (2026-06-06).** Tal como o
> `Calendars.ReadWrite` na Fase 2, as operações de Teams precisam de **admin consent** dos scopes
> `Chat.Read`/`Chat.ReadWrite` no tenant Entra. **Estado:** o admin consent foi **concedido no
> Entra em 2026-06-06**. Os scopes já estão no default de `GRAPH_SCOPES` (`config.py`) e no
> `.env.example`. **Falta ainda (operacional, na VPS):** alinhar o `GRAPH_SCOPES` do `.env` de
> produção — que se sobrepõe ao default do `config.py` (lição da Fase 2, correção pós-deploy nº 1)
> — fazer deploy da build da Fase 3 e **reiniciar**, e depois **re-login** dos utilizadores para o
> token Graph passar a incluir os scopes de Teams. Os testes mockados nunca foram bloqueados por
> isto (254 a passar).

> **Confirmação das permissões do Graph (2026-06-06).** Verificado na documentação oficial do
> Microsoft Graph que **`Chat.ReadWrite` é suficiente para criar conversa** (`POST /chats`) — a
> permissão `Chat.Create` que aparece no Entra é a alternativa de *menor privilégio* (só criar) e
> **não é necessária** além do `Chat.ReadWrite`, que já engloba criar+ler+enviar. A doc confirma
> ainda a **idempotência do 1:1** ("só pode existir uma conversa 1:1 entre dois membros; se já
> existir, devolve a existente e não cria nova") — exatamente o pressuposto de `create_one_on_one_chat`
> e do fluxo procurar→criar da US-3.4 (D1/D3). **Limite conhecido (não bug):** *descobrir* uma
> pessoa nova POR NOME depende do `resolve_recipient` (People `People.Read` + Contactos
> `Contacts.Read`), que cobre contactos e pessoas relevantes; um estranho total no diretório, sem
> qualquer relação, pode não ser resolvido (não há scope de diretório `User.ReadBasic.All` — fora
> do âmbito da Fase 3). Nesse caso, indicar o **email exato** permite criar a conversa na mesma.
> Fontes: [Create chat](https://learn.microsoft.com/en-us/graph/api/chat-post?view=graph-rest-1.0),
> [Send message in a chat](https://learn.microsoft.com/en-us/graph/api/chat-post-messages?view=graph-rest-1.0),
> [Permissions reference](https://learn.microsoft.com/en-us/graph/permissions-reference).

## Legenda

- ✅ feito · ⬜ pendente
- **Testado (automático):** coberto por testes unit/integração com Graph/Entra mockados
  (FakeGraphClient estendido).
- **Validação manual (tenant real):** execução no tenant/VPS reais — responsabilidade do
  cliente (ver runbook). Pendente em **todas** as US enquanto não houver admin consent de
  `Chat.Read`/`Chat.ReadWrite` e acesso ao tenant real.

## Tabela de estado

| US | Descrição curta | Implementado | Testado (auto) | Validação manual | Notas |
|----|-----------------|:---:|:---:|:---:|-------|
| US-3.1 | Listar chats (1:1 e grupo, filtro client-side, preview sanitizado) | ✅ | ✅ | ⬜ | `GET /me/chats?$expand=members,lastMessagePreview&$top=N` via `list_chats`. **Filtro client-side (D2)**: `filter_text` filtra em memória (case-insensitive, substring) por **tópico** (grupo) OU **nome/email** de qualquer membro — nunca `$filter`/`$search` no Graph. Só pagina (`list_chats_next`) **quando há filtro**, até satisfazer ou atingir `_MAX_LIST_FETCH=200` → sinaliza `truncated_at` (não trunca em silêncio). `members` traz só **nome+email+aad_user_id** (minimização RGPD). `last_message_preview` sanitizado (`sanitize_html`, pode vir `None` — R6) + `content_is_untrusted=true`. Ordena por `last_updated` desc (defensivo). Sem conta → `reauth_required`, Graph não tocado. |
| US-3.2 | Ler mensagens (N mais recentes, has_more, mais antigas a pedido, is_system) | ✅ | ✅ | ⬜ | `GET /me/chats/{id}/messages?$top=N&$orderby=createdDateTime desc`. `top` fixado em `[1, _MAX_MESSAGES_PER_CALL=50]`, default `25` (D4). **NÃO auto-pagina (D5)**: devolve as N mais recentes + `has_more`/`next_link`; mensagens mais antigas só a pedido via `page_token` → `list_chat_messages_next` (e NÃO `list_chat_messages`). `is_system` marcado por `messageType != "message"` (D8) — entradas/saídas, mudança de tópico, nunca acionáveis. Corpo HTML sanitizado na tool; `content_is_untrusted=true`. `chat_id` em falta → `error` sem tocar no Graph. `reauth_required` coberto. |
| US-3.3 | Enviar mensagem (prepare declara chat + N + domínios + formato; **NÃO envia**) | ✅ | ✅ | ⬜ | `teams_send_message_prepare`/`_confirm`. Validação **antes de qualquer IO**: `chat_id`/`body` obrigatórios; `body_type ∈ {text,html}` (D6, default `text`); `len(body) <= _MAX_BODY_CHARS=28000` senão `error` orientador (D10, sem token). **prepare lê via `get_chat(chat_id)` (A2)** para o resumo "Enviar mensagem no chat <1:1\|de grupo> com N participante(s) (domínios: …) [formato: …]" e **NÃO envia** (`send_chat_message` a 0). Leitura `get_chat` é best-effort (não passa por `call_graph`): se falhar por não-auth, degrada o resumo ("Detalhes do chat indisponíveis.") mas ainda emite token. confirm envia 1× e audita `teams.send` com `extra={chat_type, body_type}` — **sem `subject_hash` em extra** (A1). Replay → `idempotent_replay=true`, envio fica a 1. `reauth_required` em ambas as fases (token não consumido no confirm). |
| US-3.4 | Iniciar conversa 1:1 (obter/criar; existente → ok sem token; novo → prepare/confirm) | ✅ | ✅ | ⬜ | `teams_get_or_create_one_on_one_chat_prepare`/`_confirm`. **D3 — procurar primeiro, criar depois**: resolve o próprio email (`_own_email` via `/me`), lê os chats existentes (`list_chats`, paginando até `_MAX_LIST_FETCH`), procura um `oneOnOne` cujo único OUTRO membro (excluído o próprio) == `member_email` (case-insensitive). **Encontrado →** `{status:"ok", chat_id, is_new_chat:false}` **SEM token** (nada a criar). **Não encontrado** (inclui membro só com `userId` sem email — segue para criação) **→** `pending_confirmation` com resumo "Vai INICIAR uma nova conversa de Teams (1:1) com <email>" e token (criar = escrita, **D1**). confirm faz `create_one_on_one_chat` (idempotente no Graph) 1× e audita `teams.chat_create` (`{chat_type:"oneOnOne", is_new_chat:true}`). `member_email` em falta → `error`. Replay idempotente (criação fica a 1). `reauth` em ambas as fases. |
| US-3.5 | Responder numa conversa (= enviar no mesmo chat_id) | ✅ | ✅ | ⬜ | Em chats **não há thread/reply server-side** (D7): "responder" = enviar nova mensagem no mesmo `chat_id`. **Reusa o par `teams_send_message_prepare`/`_confirm`** de US-3.3 — sem novos invariantes nem tools. `@menções`/`reply_to_message_id` ficam FORA da v1 (diferidos, D7/D11). |

> **Validação manual pendente em TODAS as US (⬜).** Ainda não há admin consent de
> `Chat.Read`/`Chat.ReadWrite` no tenant nem validação no tenant real. As colunas Implementado
> e Testado (auto) estão ✅ em todas (254 testes a passar). Atualizar a coluna **Validação
> manual** à medida que cada US for confirmada no tenant real (ver runbook §8).

## Detalhe por user story

### US-3.1 — `run_teams_list_chats` (read)

Lê a 1ª página de `/me/chats` com `$expand=members,lastMessagePreview&$top=N` via `list_chats`.
Quando há `filter_text` **e** `next`, pagina via `list_chats_next` até satisfazer o filtro ou
atingir `_MAX_LIST_FETCH=200` (sem filtro, a 1ª página chega — não paginamos). O filtro é
**client-side (D2)**: em memória, case-insensitive substring, por `topic` OU `name`/`email` de
qualquer membro. Cada chat tem o `last_message_preview` sanitizado por `sanitize_html`
(conteúdo NÃO-confiável, pode ser `None` — R6) e a lista é ordenada por `last_updated` desc
(defensivo, caso o Graph não ordene). Resposta:
`{status:"ok", chats, count, has_more, content_is_untrusted:true}` (+ `truncated_at` se houve
truncagem). Sem conta ligada → `reauth_required`, o Graph não é tocado.

### US-3.2 — `run_teams_read_messages` (read)

`chat_id` obrigatório (`error` se faltar, sem tocar no Graph). `top` fixado em
`min(max(top,1), _MAX_MESSAGES_PER_CALL=50)`, default `25` (D4). Sem `page_token` →
`list_chat_messages(chat_id, top=top)`; com `page_token` → `list_chat_messages_next(page_token)`
(mensagens mais antigas, a pedido explícito — D5, **NÃO auto-pagina**). Cada mensagem já vem com
`is_system` marcado (`messageType != "message"`, D8); o corpo HTML é sanitizado na tool. Resposta:
`{status:"ok", chat_id, messages, count, has_more, next_link, content_is_untrusted:true}`.

### US-3.3 / US-3.5 — `run_teams_send_message_prepare` / `_confirm` (write)

Validação antes de qualquer IO: `chat_id`/`body` obrigatórios; `body_type` em `{text, html}`
(D6, default `text`); `len(body) <= _MAX_BODY_CHARS=28000` (D10, senão `error` orientador "Divida
em partes", **sem token**). O prepare resolve o token (`resolve_access_token`) e **lê o chat via
`get_chat(chat_id)` (A2)** — leitura pontual, robusta a utilizadores com >50 chats, em vez de um
match sobre a 1ª página de `list_chats`. Esta leitura é **best-effort por desenho**: não passa por
`call_graph`, pelo que uma falha não-auth (`UpstreamAuthError`/`ReauthRequired`) é apanhada e o
resumo degrada ("Detalhes do chat indisponíveis.") sem nunca marcar a conta como expirada nem
escrever (mesmo cuidado do `_resolve_tz` da Fase 2). O resumo declara o tipo de chat
(`_chat_type_label`: "1:1"/"de grupo"), N participantes e domínios (`_domains`) + formato. O
confirm chama `send_chat_message` e audita `teams.send`. **US-3.5 (responder) reusa este mesmo
par** no mesmo `chat_id` (D7 — não há thread em chats).

### US-3.4 — `run_teams_get_or_create_one_on_one_chat_prepare` / `_confirm` (write — D1/D3)

`member_email` obrigatório (já resolvido a montante — D9; `error` se faltar). **D3 — procurar
primeiro:** resolve o próprio email (`_own_email`, partilhado com `calendar.py` via `/me`), lê os
chats existentes (`list_chats`, paginando até `_MAX_LIST_FETCH`) e procura, via `_find_one_on_one`,
um `oneOnOne` cujo conjunto de outros membros (excluído o próprio) seja exatamente
`{member_email}` (comparação por email, case-insensitive). Encontrado → `{status:"ok", chat_id,
is_new_chat:false}` **SEM token**. Não encontrado (inclui o caso de o membro só trazer `userId` sem
email, que falha graciosamente e segue para criação) → `pending_confirmation` com token (criar =
escrita, **D1**). O confirm chama `create_one_on_one_chat` (idempotente no Graph) e audita
`teams.chat_create`.

## Garantias transversais (verificadas por testes)

- **prepare NÃO escreve (por contagem):** nas duas escritas, `send_chat_message` /
  `create_one_on_one_chat` ficam a **0** após o prepare. No prepare de envio, a única chamada
  Graph é a leitura `get_chat` (count=1, A2); no prepare de obter/criar 1:1, só leituras
  (`/me` + `list_chats`).
- **Idempotência (anti-duplicação):** replay de um token consumido devolve
  `idempotent_replay=true` **sem re-executar** — o envio/criação fica a 1 (risco real num chat).
- **Reautenticação graciosa:** falha de refresh no confirm → `reauth_required`, sem chamar o
  Graph para escrita; o token de confirmação **não é consumido** (`consumed_at is None`,
  repetível após re-login). Sem conta no prepare → `reauth_required`, Graph não tocado.
- **TTL / isolamento:** token expirado → `expired`; token de outro `subject` → `error`. Em
  ambos, a escrita real fica a 0. Isolamento estrito por `subject` (token delegado).
- **Sanitização + `content_is_untrusted`:** o `body` de cada mensagem e o
  `last_message_preview` passam por `sanitize_html` (quando HTML) e a resposta traz sempre
  `content_is_untrusted=true`.
- **`is_system` (D8):** mensagens de sistema (entradas/saídas, mudança de tópico) vêm
  `is_system=true` e `from=None`, marcadas como NÃO-acionáveis; cartões/anexos resumidos só
  como `attachments_count`.
- **Auditoria só-metadados (A1):** cada escrita emite `event=audit` com `subject_hash` (da
  *identidade*, no campo de topo — nunca sobrescrito), `action` (`teams.send`/`teams.chat_create`),
  `target` (o `chat_id`), `recipients_count` e `extra` seguro (`teams.send` →
  `{chat_type, body_type}`; `teams.chat_create` → `{chat_type:"oneOnOne", is_new_chat:true}`).
  **Nunca** o texto da mensagem, nomes ou emails em claro — provado por
  `assert "olá" not in str(audit)`.

## Onde estão os testes

- Unit: `tests/unit/test_graph_teams_client.py` — mapeadores (`_map_chat_summary`,
  `_map_chat_message`, `_chat_from`) e construção dos pedidos: `test_map_chat_summary_oneonone_sem_topico`,
  `test_map_chat_summary_grupo_com_topico_e_membro_sem_email`, `test_map_chat_message_normal_e_sistema`,
  `test_chat_from_aplicacao_e_nulo`, `test_list_chats_monta_expand_top_orderby`,
  `test_get_chat_monta_expand_members` (A2), `test_list_chat_messages_monta_top_orderby`,
  `test_send_chat_message_text_e_html`, `test_create_one_on_one_chat_monta_body`, e a paginação
  por `@odata.nextLink` absoluto (`test_list_chats_next_segue_link_absoluto`,
  `test_list_chat_messages_next_segue_link_absoluto`).
- Integração — leituras (US-3.1, US-3.2): `tests/integration/test_teams_read_e2e.py` —
  `test_list_chats_simples`, `test_list_chats_preview_sanitizado`, `test_list_chats_filtro_por_topico`,
  `test_list_chats_filtro_por_membro`, `test_list_chats_filtro_pagina_ate_satisfazer`,
  `test_list_chats_sem_conta_reauth`, `test_read_messages_default_top_e_sanitiza`,
  `test_read_messages_clamp_top_a_50`, `test_read_messages_has_more_e_next_link`,
  `test_read_messages_page_token_usa_next`, `test_read_messages_sem_chat_id_erro`,
  `test_read_messages_reauth`.
- Integração — escritas (US-3.3, US-3.4, US-3.5 + transversais):
  `tests/integration/test_teams_write_e2e.py` — `test_send_prepare_le_get_chat_nao_escreve` (A2),
  `test_send_prepare_get_chat_degrada_mas_emite_token`, `test_send_html_aceite`,
  `test_send_body_type_invalido_erro`, `test_send_body_demasiado_longo_erro`,
  `test_send_falta_campos_erro`, `test_send_confirm_envia_e_audita_sem_subject_hash_em_extra` (A1),
  `test_send_confirm_idempotente`, `test_send_confirm_reauth_nao_envia_nem_consome_token`,
  `test_send_prepare_reauth`, `test_us35_responder_reusa_send_no_mesmo_chat` (US-3.5),
  `test_get_or_create_chat_existente_ok_sem_token`, `test_get_or_create_chat_inexistente_pending`,
  `test_get_or_create_chat_membro_sem_email_cria`, `test_get_or_create_chat_confirm_cria_e_audita`,
  `test_get_or_create_chat_confirm_idempotente`, `test_get_or_create_chat_sem_member_email_erro`,
  `test_get_or_create_chat_prepare_reauth`, `test_send_token_expirado_devolve_expired`,
  `test_send_token_de_outro_subject_rejeitado`. `FakeGraphClient` estendido em
  `tests/integration/fake_graph.py`.

Correr: `python -m pytest -q` (**254 passed**) · lint: `python -m ruff check src tests`.
